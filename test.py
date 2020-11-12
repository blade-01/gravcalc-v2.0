import unittest

class TestInput(unittest.TestCase):
    def test_addition(self):
        add = 2 + 2
        self.assertEqual( add, 4)

if __name__ == '__main__':
    unittest.main()