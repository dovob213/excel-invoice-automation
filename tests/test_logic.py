import unittest
from src.logic import OrderParser, CatalogParser, PriceMatcher
from src.utils import normalize_string

class TestLogic(unittest.TestCase):
    def test_normalize(self):
        self.assertEqual(normalize_string("  Apple  "), "apple")
        self.assertEqual(normalize_string("Apple Pie"), "applepie")

    def test_matcher_exact(self):
        catalog = {
            "apple": [{"spec": "1kg", "price": 1000}, {"spec": "2kg", "price": 1800}]
        }
        matcher = PriceMatcher(catalog)
        
        price = matcher.get_price("Apple", "1kg")
        self.assertEqual(price, 1000)
        
        price = matcher.get_price("Apple", "2kg")
        self.assertEqual(price, 1800)
        
        price = matcher.get_price("Apple", "500g")
        self.assertIsNone(price)

    def test_matcher_partial_name(self):
        # "Name exact match first, if absent, try partial"
        catalog = {
            "GalaApple": [{"spec": "1kg", "price": 1200}]
        }
        matcher = PriceMatcher(catalog)
        
        # Exact match fails for "Apple", but "Apple" in "GalaApple" -> Partial?
        # My implementation: if search_term in catalog_key OR catalog_key in search_term
        
        # Test 1: Search "Apple", Catalog "GalaApple" -> "Apple" in "GalaApple" -> NO Match (Strict)
        price = matcher.get_price("Apple", "1kg")
        self.assertIsNone(price)

    def test_matcher_strategies(self):
        catalog = {
            "pork": [
                {"spec": "1kg,korean", "price": 5000},
                {"spec": "500g", "price": 3000}
            ],
            "onion": [
                {"spec": "10kg", "price": 2000}
            ]
        }
        matcher = PriceMatcher(catalog)
        
        # Strategy 3: Token/Permutation "korean, 1kg" vs "1kg,korean"
        price = matcher.get_price("pork", "korean, 1kg")
        self.assertEqual(price, 5000)
        
        # Strategy 4: Blank Spec "onion (None)" -> "onion (10kg)" (Unique)
        price = matcher.get_price("onion", None)
        self.assertEqual(price, 2000)
        
        # Strategy 4: Blank Spec "pork (None)" -> Multiple choices -> None
        price = matcher.get_price("pork", "")
        self.assertIsNone(price)
        
        # Strategy 5: Partial Spec "pork (1kg)" -> "pork (1kg,korean)"
        price = matcher.get_price("pork", "1kg")
        self.assertEqual(price, 5000)

if __name__ == '__main__':
    unittest.main()
