import unittest
from company_store import role_materials


class RoleMaterialTests(unittest.TestCase):
    def test_master_value_source_and_scope_are_preserved(self):
        raw = 'SampleCode17｜source｜campus｜2026-01-02｜https://source.example/item'
        item = role_materials({'公司': 'Sample', '岗位链接': 'https://alpha.example/job/12345', '内推码': raw}, 'master-hash')['items'][0]
        self.assertEqual(item['status'], 'available')
        self.assertEqual(item['value'], 'SampleCode17')
        self.assertEqual(item['raw'], raw)
        self.assertTrue(item['requireWhenAvailable'])
        self.assertEqual(item['evidence']['masterDataSha256'], 'master-hash')
        self.assertEqual(item['evidence']['roleUrl'], 'https://alpha.example/job/12345')

    def test_absent_code_is_distinct_from_unresolved_material(self):
        for raw in ('', '无公开内推码', '未找到公开推荐码'):
            self.assertEqual(role_materials({'内推码': raw}, 'hash')['items'][0]['status'], 'absent')
        for raw in ('待核查', '请联系人获取', 'foo bar', '甲码/乙码'):
            item = role_materials({'内推码': raw}, 'hash')['items'][0]
            self.assertEqual(item['status'], 'unresolved')
            self.assertIsNone(item['value'])


if __name__ == '__main__':
    unittest.main()
