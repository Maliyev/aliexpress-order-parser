import sys
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from parse_order import PageTree, parse_order_date


class OrderDateTests(unittest.TestCase):
    def read_date(self, html):
        parser = PageTree()
        parser.feed(html)
        return parse_order_date(parser.root)

    def test_english_date(self):
        self.assertEqual(self.read_date("<div>Order date 29 January 2024</div>"), date(2024, 1, 29))

    def test_date_without_year_uses_current_year(self):
        current_year = date.today().year
        self.assertEqual(self.read_date("<div>Order date 5 June</div>"), date(current_year, 6, 5))

    def test_russian_date(self):
        self.assertEqual(self.read_date("<div>Дата заказа 29 января 2024</div>"), date(2024, 1, 29))

    def test_azerbaijani_date(self):
        self.assertEqual(self.read_date("<div>Sifariş tarixi 29 yanvar 2024</div>"), date(2024, 1, 29))
