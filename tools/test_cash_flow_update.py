"""Parser checks for the office-PC Cash Flow Carnival update. No mailboxes are read."""

import unittest
from datetime import datetime
from zoneinfo import ZoneInfo

from cash_flow_update import bank_key, fingerprint, is_new, parse_message

AZ = ZoneInfo("America/Phoenix")
WHEN = datetime(2026, 11, 24, 9, 15, tzinfo=AZ)


class ParseTests(unittest.TestCase):
    def test_payment_received_is_kept(self):
        item = parse_message(
            "Payment received: Invoice #4533-(26-KB-Mark Kampf remodel-222)",
            "You received a payment of $4,750.00\nTransaction ID ABC123XYZ\nInvoice #4533",
            WHEN,
            "GC",
        )
        self.assertIsNotNone(item)
        self.assertEqual(item["invoice"], "4533")
        self.assertEqual(item["amount_cents"], 475000)
        self.assertEqual(item["qb_id"], "ABC123XYZ")
        self.assertIn("222", item["subject"])

    def test_money_on_its_way_is_ignored(self):
        self.assertIsNone(parse_message(
            "Your money's on its way",
            "Payment received $4,750.00 and $1,325.00 will be deposited",
            WHEN,
            "GC",
        ))

    def test_pending_avidpay_is_ignored(self):
        self.assertIsNone(parse_message(
            "AvidPay deposit alert",
            "AvidPay CK6531 for $50,000.00 is pending and will arrive",
            WHEN,
            "GC",
        ))

    def test_posted_bank_deposit_is_kept(self):
        item = parse_message(
            "Bank of America deposit posted",
            "A posted deposit of $50,000.00 from AvidPay CK6531",
            WHEN,
            "GC",
        )
        self.assertIsNotNone(item)
        self.assertEqual(item["kind"], "bank")
        self.assertEqual(item["amount_cents"], 5_000_000)

    def test_duplicate_quickbooks_id_is_not_new(self):
        item = parse_message(
            "Payment received: Invoice #4534-(Mark Kampf-222)",
            "Payment of $1,325.00\nTransaction ID SAMEID1",
            WHEN,
            "GC",
        )
        self.assertFalse(is_new(item, {"SAMEID1"}, set(), set()))

    def test_bank_dedupe_uses_date_amount_and_customer(self):
        item = parse_message(
            "Deposit posted",
            "Customer: Argo Marina\nA posted deposit of $50,000.00 AvidPay",
            WHEN,
            "GC",
        )
        self.assertFalse(is_new(item, set(), set(), {bank_key(item)}))
        self.assertEqual(fingerprint(item)[0], "GC")


if __name__ == "__main__":
    unittest.main()
