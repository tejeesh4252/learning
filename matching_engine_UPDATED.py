"""
matching_engine.py

Original 5 rules, plus a new Rule 6 confirmed by the business team on
2026-09-02: for aggregation-account bank disbursements where individual
lender-to-payee pairing is genuinely not possible (no allocation sheet
exists), use a payee-to-loan mapping table to group bank transactions by
which loan they relate to, and confirm the GROUP TOTAL matches the sum of
Yardi lender-level entries for that loan/date — without fabricating
individual line-item pairs that the data doesn't actually support.
"""
from __future__ import annotations
import uuid
import itertools
import datetime as dt
from collections import defaultdict
from typing import Dict, List, Any, Optional, Tuple

from src.utils.config_loader import ConfigLoader
from src.database.schema import (
    get_session, BankTransaction, YardiTransaction, Match, ReconciliationRun,
)

AMOUNT_TOLERANCE = 0.01
MAX_COMBINATION_SIZE = 3


class MatchingEngine:
    def __init__(self, config: ConfigLoader):
        self.config = config
        self.rules = sorted(self.config.get_matching_rules(), key=lambda r: r.get("priority", 99))
        self.account_map = self._build_account_map()

    def _build_account_map(self) -> Dict[str, str]:
        mapping = {}
        for entry in self.config.get_account_crosswalk():
            bank_no = entry.get("bank_account_no")
            investment = entry.get("yardi_investment")
            if bank_no and investment:
                mapping[str(bank_no).strip()] = investment
        return mapping

    def run(self) -> Dict[str, Any]:
        session = get_session(self.config.get_db_path())
        run_record = ReconciliationRun(
            run_date=dt.date.today(), run_type="manual", status="running",
            started_at=dt.datetime.utcnow(),
        )
        session.add(run_record)
        session.commit()

        try:
            bank_txns = session.query(BankTransaction).filter(
                BankTransaction.match_status == "unmatched"
            ).all()
            yardi_txns = session.query(YardiTransaction).filter(
                YardiTransaction.match_status == "unmatched"
            ).all()
            run_record.bank_transactions_considered = len(bank_txns)
            run_record.yardi_transactions_considered = len(yardi_txns)

            groups = self._group_by_account(bank_txns, yardi_txns)
            matches_made = 0
            used_bank_ids: set = set()
            used_yardi_ids: set = set()

            for account_key, (b_list, y_list) in groups.items():
                b_pool = [b for b in b_list if b.id not in used_bank_ids]
                y_pool = [y for y in y_list if y.id not in used_yardi_ids]

                for rule in self.rules:
                    if rule.get("id") == 6:
                        continue  # Rule 6 runs separately, see below
                    if not b_pool or not y_pool:
                        break
                    new_matches, b_pool, y_pool = self._apply_rule(
                        rule, b_pool, y_pool, run_record.id, session
                    )
                    matches_made += len(new_matches)
                    for m in new_matches:
                        if m.bank_transaction_id:
                            used_bank_ids.add(m.bank_transaction_id)
                        if m.yardi_transaction_id:
                            used_yardi_ids.add(m.yardi_transaction_id)

            session.commit()

            # --- Rule 6 runs separately from the normal per-account groups
            # above, because aggregation accounts deliberately have NO
            # single investment mapping (that's the whole point of them) —
            # they can't be pre-grouped the way a normal fund account can.
            # Rule 6 does its own investment lookup via the payee mapping.
            agg_accounts = self._get_aggregation_account_numbers()
            if agg_accounts:
                remaining_bank = session.query(BankTransaction).filter(
                    BankTransaction.match_status == "unmatched",
                    BankTransaction.account_no.in_(agg_accounts),
                ).all()
                remaining_yardi = session.query(YardiTransaction).filter(
                    YardiTransaction.match_status == "unmatched",
                ).all()
                rule6 = next((r for r in self.rules if r.get("id") == 6), None)
                if rule6 and remaining_bank and remaining_yardi:
                    new_matches, _, _ = self._apply_rule(
                        rule6, remaining_bank, remaining_yardi, run_record.id, session
                    )
                    matches_made += len(new_matches)
                    session.commit()

            run_record.matches_found = matches_made
            run_record.status = "completed"
            run_record.completed_at = dt.datetime.utcnow()
            session.commit()

            return {
                "run_id": run_record.id,
                "bank_considered": run_record.bank_transactions_considered,
                "yardi_considered": run_record.yardi_transactions_considered,
                "matches_found": matches_made,
                "status": "completed",
            }
        except Exception as e:
            session.rollback()
            run_record.status = "failed"
            run_record.notes = str(e)
            session.commit()
            return {"run_id": run_record.id, "status": "failed", "error": str(e)}
        finally:
            session.close()

    def _get_aggregation_account_numbers(self):
        return {
            str(entry.get("bank_account_no")).strip()
            for entry in self.config.get_account_crosswalk()
            if entry.get("account_type") and "aggregation" in entry.get("account_type", "").lower()
            and entry.get("bank_account_no")
        }

    def _group_by_account(self, bank_txns, yardi_txns):
        groups: Dict[str, Tuple[List, List]] = {}
        for b in bank_txns:
            key = str(b.account_no).strip() if b.account_no else None
            if key is None or key not in self.account_map:
                continue
            groups.setdefault(key, ([], []))[0].append(b)
        for y in yardi_txns:
            matching_bank_no = None
            for bank_no, investment in self.account_map.items():
                if investment and y.investment and investment.strip() == y.investment.strip():
                    matching_bank_no = bank_no
                    break
            if matching_bank_no is None:
                continue
            groups.setdefault(matching_bank_no, ([], []))[1].append(y)
        return groups

    def _apply_rule(self, rule, b_pool, y_pool, run_id, session):
        rule_id = rule.get("id")
        rule_name = rule.get("name", f"Rule {rule_id}")
        tolerance_days = rule.get("tolerance_days")
        match_on = rule.get("match_on", [])

        matches: List[Match] = []
        if match_on == ["tran_date", "amount"]:
            matches = self._match_date_amount(b_pool, y_pool, tolerance_days or 0, rule_id, rule_name, run_id)
        elif match_on == ["reference", "amount"]:
            matches = self._match_reference_amount(b_pool, y_pool, rule_id, rule_name, run_id)
        elif match_on == ["amount_sum"] and rule_id == 4:
            matches = self._match_one_to_many(b_pool, y_pool, tolerance_days or 0, rule_id, rule_name, run_id)
        elif match_on == ["amount_sum"] and rule_id == 5:
            matches = self._match_many_to_one(b_pool, y_pool, tolerance_days or 0, rule_id, rule_name, run_id)
        elif match_on == ["payee_group"]:
            matches = self._match_payee_group(b_pool, y_pool, rule_id, rule_name, run_id)

        matched_b_ids, matched_y_ids = set(), set()
        for m in matches:
            session.add(m)
            if m.bank_transaction_id:
                matched_b_ids.add(m.bank_transaction_id)
                bt = session.get(BankTransaction, m.bank_transaction_id)
                if bt:
                    bt.match_status = "matched"
            if m.yardi_transaction_id:
                matched_y_ids.add(m.yardi_transaction_id)
                yt = session.get(YardiTransaction, m.yardi_transaction_id)
                if yt:
                    yt.match_status = "matched"
        session.flush()

        remaining_b = [b for b in b_pool if b.id not in matched_b_ids]
        remaining_y = [y for y in y_pool if y.id not in matched_y_ids]
        return matches, remaining_b, remaining_y

    def _match_date_amount(self, b_pool, y_pool, tolerance_days, rule_id, rule_name, run_id):
        matches = []
        used_y = set()
        for b in b_pool:
            if b.amount is None or b.tran_date is None:
                continue
            for y in y_pool:
                if y.id in used_y or y.total_transaction is None or y.tran_date is None:
                    continue
                if abs(abs(b.amount) - abs(y.total_transaction)) > AMOUNT_TOLERANCE:
                    continue
                day_diff = abs((b.tran_date - y.tran_date).days)
                if day_diff > tolerance_days:
                    continue
                matches.append(self._build_match(
                    run_id, rule_id, rule_name, b, y, abs(b.amount),
                    confidence=100.0 if tolerance_days == 0 else max(50.0, 100.0 - day_diff * 10)
                ))
                used_y.add(y.id)
                break
        return matches

    def _match_reference_amount(self, b_pool, y_pool, rule_id, rule_name, run_id):
        matches = []
        used_y = set()
        for b in b_pool:
            ref = (b.serial or "").strip()
            if not ref or b.amount is None:
                continue
            for y in y_pool:
                if y.id in used_y:
                    continue
                control = (y.control_number or "").strip()
                if not control or y.total_transaction is None:
                    continue
                if control == ref and abs(abs(b.amount) - abs(y.total_transaction)) <= AMOUNT_TOLERANCE:
                    matches.append(self._build_match(run_id, rule_id, rule_name, b, y, abs(b.amount), confidence=95.0))
                    used_y.add(y.id)
                    break
        return matches

    def _match_one_to_many(self, b_pool, y_pool, tolerance_days, rule_id, rule_name, run_id):
        matches = []
        used_y = set()
        for b in b_pool:
            if b.amount is None or b.tran_date is None:
                continue
            candidates = [
                y for y in y_pool if y.id not in used_y and y.total_transaction is not None
                and y.tran_date is not None and abs((b.tran_date - y.tran_date).days) <= tolerance_days
            ]
            found = self._find_subset_sum(candidates, abs(b.amount), key=lambda y: y.total_transaction)
            if found:
                group_id = str(uuid.uuid4())
                for y in found:
                    matches.append(self._build_match(run_id, rule_id, rule_name, b, y, abs(b.amount), confidence=80.0, group_id=group_id))
                    used_y.add(y.id)
        return matches

    def _match_many_to_one(self, b_pool, y_pool, tolerance_days, rule_id, rule_name, run_id):
        matches = []
        used_b = set()
        for y in y_pool:
            if y.total_transaction is None or y.tran_date is None:
                continue
            candidates = [
                b for b in b_pool if b.id not in used_b and b.amount is not None
                and b.tran_date is not None and abs((b.tran_date - y.tran_date).days) <= tolerance_days
            ]
            found = self._find_subset_sum(candidates, abs(y.total_transaction), key=lambda b: b.amount)
            if found:
                group_id = str(uuid.uuid4())
                for b in found:
                    matches.append(self._build_match(run_id, rule_id, rule_name, b, y, abs(y.total_transaction), confidence=80.0, group_id=group_id))
                    used_b.add(b.id)
        return matches

    # ------------------------------------------------------------------
    # NEW Rule 6: Payee-mapped group total match
    # ------------------------------------------------------------------
    def _match_payee_group(self, b_pool, y_pool, rule_id, rule_name, run_id):
        """
        For bank transactions whose narrative mentions a known payee (per
        config's payee_loan_mapping), group them by (investment, date) and
        compare the GROUP TOTAL against the sum of Yardi lender-level
        entries for that investment/date. Confirmed by the business team:
        there is currently no data linking individual lender contributions
        to individual payees, so this deliberately does NOT create
        line-by-line pairs — only a validated group-total match, with every
        transaction in the group linked via the same match_group_id so a
        reviewer can see the whole picture together.
        """
        payee_map = {
            e["payee_name"].strip().lower(): e["investment"]
            for e in self.config.get_payee_loan_mapping()
            if e.get("payee_name") and e.get("investment")
        }
        if not payee_map:
            return []

        matches = []
        used_b, used_y = set(), set()
        bank_groups: Dict[Tuple[str, Any], List[BankTransaction]] = defaultdict(list)

        for b in b_pool:
            if not b.narrative or b.tran_date is None:
                continue
            narrative_lower = b.narrative.lower()
            for payee_name, investment in payee_map.items():
                if payee_name in narrative_lower:
                    bank_groups[(investment, b.tran_date)].append(b)
                    break

        for (investment, tran_date), b_group in bank_groups.items():
            if any(b.id in used_b for b in b_group):
                continue
            y_group = [
                y for y in y_pool
                if y.investment == investment and y.tran_date == tran_date and y.id not in used_y
            ]
            if not y_group:
                continue

            b_total = sum(abs(b.amount) for b in b_group)
            y_total = sum(abs(y.total_transaction) for y in y_group)

            if abs(b_total - y_total) <= AMOUNT_TOLERANCE:
                group_id = str(uuid.uuid4())
                note = (
                    f"Group total match via payee-to-loan mapping: "
                    f"{len(b_group)} bank transaction(s) totaling {b_total:,.2f} "
                    f"= {len(y_group)} Yardi lender entries totaling {y_total:,.2f}. "
                    f"Individual line-item pairing is not available in source data "
                    f"(confirmed 2026-09-02) — this is a validated GROUP match only."
                )
                for b in b_group:
                    m = self._build_match(run_id, rule_id, rule_name, b, None, abs(b.amount), confidence=70.0, group_id=group_id)
                    m.notes = note
                    matches.append(m)
                    used_b.add(b.id)
                for y in y_group:
                    m = self._build_match(run_id, rule_id, rule_name, None, y, abs(y.total_transaction), confidence=70.0, group_id=group_id)
                    m.notes = note
                    matches.append(m)
                    used_y.add(y.id)
        return matches

    @staticmethod
    def _find_subset_sum(items, target, key):
        if len(items) < 2:
            return None
        for size in range(2, min(MAX_COMBINATION_SIZE, len(items)) + 1):
            for combo in itertools.combinations(items, size):
                total = sum(abs(key(i)) for i in combo)
                if abs(total - target) <= AMOUNT_TOLERANCE:
                    return list(combo)
        return None

    @staticmethod
    def _build_match(run_id, rule_id, rule_name, bank_txn, yardi_txn, amount, confidence, group_id=None):
        return Match(
            reconciliation_run_id=run_id,
            match_group_id=group_id or str(uuid.uuid4()),
            bank_transaction_id=bank_txn.id if bank_txn else None,
            yardi_transaction_id=yardi_txn.id if yardi_txn else None,
            matched_rule_id=rule_id, matched_rule_name=rule_name,
            matched_amount=amount, confidence_score=confidence,
            status="matched", matched_at=dt.datetime.utcnow(),
        )
