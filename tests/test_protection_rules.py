"""Regression tests for the paragraph-protection rules.

The public release replaced a literal underwriter name inside the
"独立核查程序与结论保留" rule with a generic securities-firm pattern. These tests
pin the rule's trigger surface so a future sanitiser pass cannot quietly narrow
it again, and so a future "generalisation" cannot quietly widen it either.
"""

import re
import unittest

from workbench.model import INDEPENDENT_REVIEW, SECURITIES_FIRM


class SecuritiesFirmPatternTests(unittest.TestCase):
    """The firm fragment must require the word 证券, not match any characters."""

    def test_requires_the_securities_word(self):
        self.assertNotIn(r"\S{2,10}", SECURITIES_FIRM)
        self.assertIn("证券", SECURITIES_FIRM)

    def test_matches_plain_and_suffixed_firm_names(self):
        for name in ("广发证券", "中信证券股份有限公司", "某某证券有限责任公司", "示例证券"):
            with self.subTest(name=name):
                self.assertIsNotNone(re.search(SECURITIES_FIRM, name))

    def test_does_not_match_organisations_without_the_word(self):
        for name in ("发行人", "会计师事务所", "律师事务所", "证券市场"):
            with self.subTest(name=name):
                self.assertIsNone(re.search(SECURITIES_FIRM, name))


class IndependentReviewRuleTests(unittest.TestCase):
    """应命中：各种证券公司写法（含持股/有限后缀）。"""

    MATCH = (
        "经广发证券对发行人本次债券发行相关事项进行核查",
        "经中信证券股份有限公司对发行人主体资格认为符合规定",
        "经某某证券有限责任公司对发行人合法有效性进行核查",
        "经示例证券对募集资金用途核查无误",
    )

    """不应命中：不含「证券」二字的机构，或不是「经<机构>对」结构。"""

    NO_MATCH = (
        "经发行人对本次债券发行事项进行核查",
        "经会计师事务所对财务报表进行核查",
        "经律师事务所对法律事项进行核查",
        "证券市场对本次发行的影响分析",
        "经项目组对本次债券发行核查",
    )

    def test_matches_securities_firm_statements(self):
        for text in self.MATCH:
            with self.subTest(text=text):
                self.assertIsNotNone(INDEPENDENT_REVIEW.match(text), text)

    def test_ignores_non_securities_organisations(self):
        for text in self.NO_MATCH:
            with self.subTest(text=text):
                self.assertIsNone(INDEPENDENT_REVIEW.match(text), text)


class IndependentReviewRuleBoundaryTests(unittest.TestCase):
    """其它分支未被这次泛化改动波及。"""

    STILL_MATCH = (
        "经项目组查阅发行人相关资料",
        "经项目组核对相关情况",
        "经查阅审计报告后认为",
        "经查阅律师执业许可证",
        "项目组已核查相关事项",
    )

    def test_other_branches_survive(self):
        for text in self.STILL_MATCH:
            with self.subTest(text=text):
                self.assertIsNotNone(INDEPENDENT_REVIEW.match(text), text)


if __name__ == "__main__":
    unittest.main()
