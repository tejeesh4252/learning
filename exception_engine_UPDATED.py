"""
exception_engine.py

Phase 5 exception classification, now extended with the aggregation-account
net-zero logic confirmed in the 2026-09-02 business requirements call.
"""
from __future__ import annotations
import datetime as dt
from collections import defaultdict
from typing import Dict, Any, List, Set, Tuple

from src.utils.config_loader import ConfigLoader
from src.database.schema import (
    get_session, BankTransaction, YardiTransaction, ReconciliationException,
)

AMOUNT_TOLERANCE = 0.01

SEVERITY_DEFAULT = {
    "cash_received_not_booked": "High",
    "booked_not_received": "High",
    "duplicate": "Medium",
    "aggregation_variance": "Critical",
    "aggregation_surplus_no_yardi_link": "Low",
}


class ExceptionEngine:
    def __init__(self, config: ConfigLoader):
        self.config = config
        self.categories = self.config.get_exception_categories()

    def run(self, reconciliation_run_id: int = None) -> Dict[str, int]:
        session = get_session(self.config.get_db_path())
        counts: Dict[str, int] = defaultdict(int)
        try:
            session.query(ReconciliationException).filter(
                ReconciliationException.status == "Open"
            ).delete()
            session.commit()

            # --- Aggregation net-zero check runs FIRST, since it determines
            # which unmatched bank transactions get suppressed from the
            # normal "cash_received_not_booked" flag below. ---
            agg_count, suppressed_bank_ids = self._flag_aggregation_surplus(
                session, reconciliation_run_id
            )
            counts["aggregation_surplus_no_yardi_link"] = agg_count

            counts["cash_received_not_booked"] = self._flag_cash_received_not_booked(
                session, reconciliation_run_id, exclude_ids=suppressed_bank_ids
            )
            counts["booked_not_received"] = self._flag_booked_not_received(
                session, reconciliation_run_id
            )
            counts["duplicate"] = self._flag_duplicates(
                session, reconciliation_run_id, exclude_ids=suppressed_bank_ids
            )

            session.commit()
            return dict(counts)
        finally:
            session.close()

    def _severity_for(self, category: str) -> str:
        configured = self.categories.get(category, {}).get("severity")
        return configured or SEVERITY_DEFAULT.get(category, "Medium")

    def _description_for(self, category: str) -> str:
        return self.categories.get(category, {}).get("description", category)

    # ------------------------------------------------------------------
    # NEW: Aggregation account net-zero group reconciliation
    # ------------------------------------------------------------------
    def _get_aggregation_account_numbers(self) -> Set[str]:
        return {
            str(entry.get("bank_account_no")).strip()
            for entry in self.config.get_account_crosswalk()
            if entry.get("account_type") and "aggregation" in entry.get("account_type", "").lower()
            and entry.get("bank_account_no")
        }

    def _flag_aggregation_surplus(self, session, run_id) -> Tuple[int, Set[int]]:
        """
        Confirmed by the business team (2026-09-02 minutes): the Aggregation
        Account has NO fund association in Yardi for ANY transaction type —
        it's a pure pass-through. Money in should equal money out on any
        given day. Some legs (e.g. repayment routed back to individual
        funds) DO have Yardi entries and get matched by the normal rules;
        others (e.g. surplus repatriated straight to the borrower) never
        get a Yardi entry at all, by design.

        Rather than flag every unmatched leg as "missing from Yardi", this
        checks whether a day's FULL bank activity on the account (matched
        + unmatched together) nets to zero. If it does, whatever's left
        unmatched is legitimate account mechanics — logged as a Low-severity
        informational item, not a real exception. If the group does NOT
        net to zero, something is genuinely missing, and normal exception
        flagging applies unchanged.
        """
        agg_accounts = self._get_aggregation_account_numbers()
        if not agg_accounts:
            return 0, set()

        suppressed_ids: Set[int] = set()
        flagged = 0

        for acct_no in agg_accounts:
            rows = session.query(BankTransaction).filter(
                BankTransaction.account_no == acct_no
            ).all()
            by_date: Dict[Any, List[BankTransaction]] = defaultdict(list)
            for b in rows:
                if b.tran_date:
                    by_date[b.tran_date].append(b)

            for tran_date, group in by_date.items():
                unmatched_in_group = [b for b in group if b.match_status == "unmatched"]
                if not unmatched_in_group:
                    continue

                total_net = sum(b.amount or 0.0 for b in group)
                if abs(total_net) <= AMOUNT_TOLERANCE:
                    for b in unmatched_in_group:
                        session.add(ReconciliationException(
                            reconciliation_run_id=run_id,
                            category="aggregation_surplus_no_yardi_link",
                            severity=self._severity_for("aggregation_surplus_no_yardi_link"),
                            account_no=b.account_no,
                            bank_transaction_id=b.id,
                            amount=b.amount,
                            tran_date=b.tran_date,
                            description=(
                                f"{self._description_for('aggregation_surplus_no_yardi_link')}. "
                                f"Full account activity for {tran_date} nets to "
                                f"{total_net:,.2f} (effectively zero) — this leg "
                                f"(amount {b.amount:,.2f}) has no Yardi counterpart "
                                f"by design."
                            ),
                        ))
                        suppressed_ids.add(b.id)
                        flagged += 1
        return flagged, suppressed_ids

    # ------------------------------------------------------------------
    def _flag_cash_received_not_booked(self, session, run_id, exclude_ids: Set[int] = None) -> int:
        exclude_ids = exclude_ids or set()
        rows = session.query(BankTransaction).filter(
            BankTransaction.match_status == "unmatched",
            BankTransaction.amount > 0,
        ).all()
        rows = [b for b in rows if b.id not in exclude_ids]
        for b in rows:
            session.add(ReconciliationException(
                reconciliation_run_id=run_id,
                category="cash_received_not_booked",
                severity=self._severity_for("cash_received_not_booked"),
                account_no=b.account_no,
                bank_transaction_id=b.id,
                amount=b.amount,
                tran_date=b.tran_date,
                description=(
                    f"{self._description_for('cash_received_not_booked')}. "
                    f"Bank txn {b.serial or b.id}, amount {b.amount:,.2f} on {b.tran_date}."
                ),
            ))
        return len(rows)

    def _flag_booked_not_received(self, session, run_id) -> int:
        rows = session.query(YardiTransaction).filter(
            YardiTransaction.match_status == "unmatched",
        ).all()
        for y in rows:
            session.add(ReconciliationException(
                reconciliation_run_id=run_id,
                category="booked_not_received",
                severity=self._severity_for("booked_not_received"),
                account_no=None,
                yardi_transaction_id=y.id,
                amount=y.total_transaction,
                tran_date=y.tran_date,
                description=(
                    f"{self._description_for('booked_not_received')}. "
                    f"Yardi investment '{y.investment}', amount "
                    f"{(y.total_transaction or 0):,.2f} on {y.tran_date}."
                ),
            ))
        return len(rows)

    def _flag_duplicates(self, session, run_id, exclude_ids: Set[int] = None) -> int:
        exclude_ids = exclude_ids or set()
        rows = session.query(BankTransaction).filter(
            BankTransaction.match_status == "unmatched",
        ).all()
        rows = [b for b in rows if b.id not in exclude_ids]
        groups: Dict[tuple, List[BankTransaction]] = defaultdict(list)
        for b in rows:
            if b.amount is None or b.tran_date is None:
                continue
            key = (b.account_no, b.tran_date, round(b.amount, 2))
            groups[key].append(b)

        flagged = 0
        for key, items in groups.items():
            if len(items) < 2:
                continue
            for b in items:
                session.add(ReconciliationException(
                    reconciliation_run_id=run_id,
                    category="duplicate",
                    severity=self._severity_for("duplicate"),
                    account_no=b.account_no,
                    bank_transaction_id=b.id,
                    amount=b.amount,
                    tran_date=b.tran_date,
                    description=(
                        f"{self._description_for('duplicate')}. "
                        f"{len(items)} unmatched transactions share account "
                        f"{b.account_no}, date {b.tran_date}, amount {b.amount:,.2f}."
                    ),
                ))
                flagged += 1
        return flagged
