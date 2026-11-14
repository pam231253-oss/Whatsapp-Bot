import datetime as dt
import unittest
from unittest.mock import patch

from chama_core import Bot, split_messages


class FakeBackend:
    def __init__(self):
        self.sheets = {
            'Setup': [
                ['Umoja Chama'], ['KES'], [dt.date(2026, 11, 14)],
                [None], [None], [None], [None], [None],
                ['Paybill 123'], ['Treasurer'], ['254700000000'],
            ],
            'Members': [
                ['REG001', 'Grace Wanjiru', '254700000001', dt.date(2026, 1, 1), 1000, 100, None, None],
                ['REG002', 'John Kamau', '254700000002', dt.date(2026, 1, 1), 2000, 200, None, None],
                ['REG003', 'Mercy Achieng', '254700000003', dt.date(2026, 1, 1), 500, 100, None, None],
                ['REG004', 'Peter Otieno', '254700000004', dt.date(2026, 12, 1), 1000, 100, None, None],
            ],
            'Transactions': [
                [dt.date(2026, 11, 1), 'John Kamau', None, 'Savings', 2000],
                [dt.date(2026, 11, 2), 'Grace Wanjiru', None, 'Savings', 1000],
                [dt.date(2026, 11, 2), 'John Kamau', None, 'Welfare', 200],
                [dt.date(2026, 10, 31), 'Mercy Achieng', None, 'Savings', 500],
            ],
            'Summary': [
                ['REG001', 'Grace Wanjiru'] + [None] * 17 + [0],
                ['REG002', 'John Kamau'] + [None] * 17 + [2500],
                ['REG003', 'Mercy Achieng'] + [None] * 17 + [100],
            ],
            'Dashboard': [
                ['Total savings collected', 12000],
                ['Contribution arrears (savings + welfare + MGR)', 3100],
                ['Outstanding loan balance', 5000],
                ['of which OVERDUE', 1000],
                ['Welfare fund balance', 2500],
                ['EXPECTED CASH ON HAND', 22000],
            ],
        }

    def get(self, sheet, _a1):
        return self.sheets.get(sheet, [])


class ReportingTests(unittest.TestCase):
    def setUp(self):
        self.bot = Bot(FakeBackend(), ['254700000000'])

    @patch('chama_core.today', return_value=dt.date(2026, 11, 14))
    def test_daily_report_filters_day_and_totals_categories(self, _today):
        reply = self.bot.handle('254700000000', 'daily 2 nov')
        self.assertIn('DAILY CONTRIBUTIONS - 02 Nov 2026', reply)
        self.assertIn('Grace Wanjiru: Savings 1,000', reply)
        self.assertIn('John Kamau: Welfare 200', reply)
        self.assertIn('Total received: KES 1,200', reply)
        self.assertNotIn('Mercy Achieng', reply)

    @patch('chama_core.today', return_value=dt.date(2026, 11, 14))
    def test_monthly_report_compares_savings_with_active_members(self, _today):
        reply = self.bot.handle('254700000000', 'monthly last')
        self.assertIn('MONTHLY CONTRIBUTIONS - October 2026', reply)
        self.assertIn('Savings against expected: KES 500 / 3,500 (14%)', reply)
        self.assertIn('*No savings payment:* Grace Wanjiru, John Kamau', reply)

    @patch('chama_core.today', return_value=dt.date(2026, 11, 14))
    def test_group_statement_includes_member_status_and_forwarding_details(self, _today):
        reply = self.bot.handle('254700000000', 'group')
        self.assertIn('*UMOJA CHAMA: GROUP STATEMENT*', reply)
        self.assertIn('⚠️ John Kamau: owes 2,500', reply)
        self.assertIn('EXPECTED CASH ON HAND: KES 22,000', reply)
        self.assertIn('Paybill 123', reply)
        self.assertIn('Treasurer: Treasurer 254700000000', reply)

    def test_member_cannot_access_admin_reports(self):
        reply = self.bot.handle('254700000002', 'group last')
        self.assertIn('only show your own', reply)
        self.assertNotIn('GROUP STATEMENT', reply)

    def test_long_messages_are_split_without_exceeding_limit(self):
        text = 'first line\n' + ('word ' * 20) + '\nlast line'
        messages = split_messages(text, limit=24)
        self.assertTrue(all(len(message) <= 24 for message in messages))
        self.assertEqual(messages[0], 'first line')
        self.assertTrue(messages[-1].endswith('last line'))
        self.assertEqual(''.join(messages).replace('\n', ''), text.replace('\n', ''))


if __name__ == '__main__':
    unittest.main()
