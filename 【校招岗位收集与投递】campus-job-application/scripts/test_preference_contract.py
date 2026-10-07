"""Synthetic examples only; validation never interprets natural-language JD truth."""
import copy
import importlib.util
from pathlib import Path
import unittest
from preference_contract import validate_snapshot, identity, canonical_url

ROOT = Path(__file__).resolve().parents[1]


def preferences():
    return {'version': 'p2', 'confirmation_ref': 'synthetic-confirmation', 'effective_at': 'synthetic-effective-time', 'supersedes_version': None, 'recheck_scope': [], 'hard_scope': {'cities': []},
            'axes': {axis: {'raw': 'synthetic explicitly unrestricted', 'relations': []}
                     for axis in ['employment', 'city', 'direction']},
            'internship_policy': 'conditional', 'conflict_policy': 'user_decides'}


def role(job_id='1'):
    return {'company_key': 'synthetic-ats', 'company': 'Synthetic Employer', 'official_job_id': job_id,
            'role_key': 'synthetic-ats:id:' + job_id, 'title': '数据产品',
            'direct_url': 'https://example.test/jobs/' + job_id, 'bu': '未披露', 'cities': ['合成其他城市'],
            'employment_type': 'formal', 'direction': 'AI product',
            'jd_full_text': 'Synthetic JD: own AI product design, Agent requirements and iteration.',
            'jd_capture_ref': 'synthetic-full-jd', 'jd_basis': 'AI product design and Agent iteration',
            'decision_reason': 'User-confirmed function matches full JD', 'decision_basis': 'jd',
            'hard_scope_pass': True, 'disposition': 'primary', 'preference_version': 'p2',
            'stage': 'verified', 'pre_form_report_ref': 'synthetic-role-report', 'report_precedes_form': True,
            'execution': {action: {'status': 'not_started'} for action in ['fill', 'save', 'submit']}}


def company():
    return {'company_key': 'synthetic-ats', 'formal_coverage': 'sufficient',
            'formal_coverage_refs': ['synthetic-plans-bus-pagination-jds'],
            'suitable_formal_found': False, 'no_suitable_formal_ref': 'synthetic-screening'}


def load(path):
    spec = importlib.util.spec_from_file_location('validator_' + path.stem, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class PreferenceTests(unittest.TestCase):
    def check(self, prefs=None, roles=None, companies=None, prior=None):
        return validate_snapshot(prefs or preferences(), roles if roles is not None else [role()],
                                 companies if companies is not None else [company()], prior if prior is not None else [])

    def test_unrestricted_other_city_remains_candidate_with_city_sort(self):
        p = preferences()
        p['axes']['city']['relations'] = [{'higher': '合成中心城市', 'lower': '合成其他城市', 'relation': 'strict'}]
        self.assertEqual(self.check(p), [])

    def test_hard_city_restriction_is_separate(self):
        p = preferences(); p['hard_scope']['cities'] = ['合成中心城市']
        self.assertTrue(any('hard city scope' in e for e in self.check(p)))

    def test_equal_weak_unknown_relations_are_preserved(self):
        p = preferences()
        p['axes']['direction']['relations'] = [dict(higher='A', lower='B', relation=r) for r in ['weak', 'tied', 'unknown']]
        before = copy.deepcopy(p)
        self.assertEqual(self.check(p), []); self.assertEqual(p, before)

    def test_cross_axis_conflict_is_displayed_without_automatic_choice(self):
        r = role(); r['cross_axis_conflict'] = True
        self.assertEqual(self.check(roles=[r]), [])
        r['selected'] = True
        self.assertTrue(any('explicit user selection' in e for e in self.check(roles=[r])))
        r['selection_ref'] = 'synthetic-user-choice'
        self.assertEqual(self.check(roles=[r]), [])

    def test_weighted_conflict_choice_rejected(self):
        p = preferences(); p['combined_weights'] = {'city': 3, 'direction': 2}
        self.assertTrue(any('weighted' in e for e in self.check(p)))

    def test_ai_data_product_candidate_and_pure_warehouse_exclusion_use_jd(self):
        ai = role(); warehouse = role('2')
        warehouse.update(jd_full_text='Synthetic JD: ETL, warehouse maintenance and reports only.',
                         direction='warehouse', jd_basis='ETL and warehouse maintenance only',
                         disposition='excluded', decision_reason='Confirmed function preference excludes warehouse-only duties')
        self.assertEqual(self.check(roles=[ai, warehouse]), [])
        warehouse['decision_basis'] = 'title_blacklist'
        self.assertTrue(any('title blacklist' in e for e in self.check(roles=[warehouse])))

    def test_formal_role_prevents_internship_fallback(self):
        r = role(); r.update(employment_type='internship', disposition='fallback', conversion_ref='synthetic-conversion')
        c = company(); c['suitable_formal_found'] = True
        self.assertTrue(any('no suitable formal' in e for e in self.check(roles=[r], companies=[c])))

    def test_visible_formal_primary_contradicts_fallback_even_if_company_flag_is_wrong(self):
        intern = role('2'); intern.update(employment_type='internship', disposition='fallback', conversion_ref='synthetic')
        self.assertTrue(any('formal primary' in e for e in self.check(roles=[role(), intern])))

    def test_sufficient_no_formal_coverage_allows_separate_fallback(self):
        r = role(); r.update(employment_type='internship', disposition='fallback', conversion_ref='synthetic-conversion')
        self.assertEqual(self.check(roles=[r]), [])

    def test_incomplete_coverage_is_pending_not_no_formal(self):
        r = role(); r.update(employment_type='internship', disposition='pending')
        c = company(); c.update(formal_coverage='incomplete', suitable_formal_found=None)
        self.assertEqual(self.check(roles=[r], companies=[c]), [])
        c['no_suitable_formal_found'] = True
        self.assertTrue(any('sufficient coverage' in e for e in self.check(roles=[r], companies=[c])))
        r.update(disposition='fallback', conversion_ref='synthetic-conversion')
        self.assertTrue(self.check(roles=[r], companies=[c]))

    def test_internship_not_mixed_into_formal_primary(self):
        r = role(); r['employment_type'] = 'internship'
        self.assertTrue(any('mix into formal' in e for e in self.check(roles=[r])))
        p = preferences(); p['internship_policy'] = 'separate'
        self.assertEqual(self.check(p, roles=[r]), [])

    def test_existing_submitted_role_and_receipt_preserved(self):
        old = {'role_key': role()['role_key'], 'stage': 'submitted', 'receipt_ref': 'synthetic-receipt'}
        self.assertTrue(any('preserve submitted' in e for e in self.check(prior=[old])))
        r = role(); r.update(stage='submitted', receipt_ref='synthetic-receipt')
        r['execution']['fill'] = {'status': 'complete', 'evidence_ref': 'synthetic-fill'}
        r['execution']['submit'] = {'status': 'complete', 'evidence_ref': 'synthetic-receipt'}
        self.assertEqual(self.check(roles=[r], prior=[old]), [])
        self.assertTrue(self.check(roles=[], prior=[old]))

    def test_explicit_selection_cannot_be_removed_by_preference_update(self):
        old = {'role_key': role()['role_key'], 'selected': True}
        self.assertTrue(any('preserve explicit selection' in e for e in self.check(prior=[old])))
        r = role(); r.update(selected=True, selection_ref='synthetic-original-choice')
        self.assertEqual(self.check(roles=[r], prior=[old]), [])

    def test_old_preference_version_rejected(self):
        r = role(); r['preference_version'] = 'p1'
        self.assertTrue(any('stale role' in e for e in self.check(roles=[r])))

    def test_missing_bu_requires_explicit_undisclosed(self):
        r = role(); r['bu'] = ''
        self.assertTrue(any('missing bu' in e for e in self.check(roles=[r])))
        r['bu'] = '未披露'; self.assertEqual(self.check(roles=[r]), [])

    def test_missing_pre_form_role_report_fails(self):
        r = role(); del r['pre_form_report_ref']
        self.assertTrue(any('pre_form_report_ref' in e for e in self.check(roles=[r])))
        r = role(); r['report_precedes_form'] = False
        self.assertTrue(any('before form' in e for e in self.check(roles=[r])))

    def test_login_only_is_not_execution_evidence(self):
        r = role(); r['execution'] = {'login': {'status': 'complete'}}
        self.assertEqual(sum('actual execution state' in e for e in self.check(roles=[r])), 3)
        r = role(); r['execution']['save'] = {'status': 'complete'}
        self.assertTrue(any('needs evidence' in e for e in self.check(roles=[r])))

    def test_click_is_not_submit_receipt(self):
        r = role(); r['execution']['submit'] = {'status': 'complete', 'evidence_ref': 'click'}
        self.assertTrue(any('receipt' in e for e in self.check(roles=[r])))

    def test_top5_and_capability_score_cannot_resolve_preferences(self):
        for field in ['score_includes_preferences', 'selected_by_top5']:
            r = role(); r[field] = True
            self.assertTrue(any('Top5' in e for e in self.check(roles=[r])))

    def test_same_title_different_bu_city_ids_remain_distinct(self):
        a, b = role(), role('2'); b.update(bu='Synthetic BU', cities=['合成中心城市'])
        self.assertEqual(self.check(roles=[a, b]), [])

    def test_multi_city_role_not_double_counted(self):
        r = role(); r['cities'].append('合成中心城市')
        self.assertEqual(self.check(roles=[r]), [])
        duplicate = copy.deepcopy(r); duplicate['cities'] = ['合成中心城市']
        self.assertTrue(any('duplicate official role' in e for e in self.check(roles=[r, duplicate])))

    def test_url_identity_keeps_business_parameters(self):
        r = role(); del r['official_job_id']
        r['direct_url'] = 'https://example.test/job?bu=2&id=7&utm_campaign=x'
        r['role_key'] = identity(r)
        self.assertEqual(canonical_url(r['direct_url']), 'https://example.test/job?bu=2&id=7')
        self.assertEqual(self.check(roles=[r]), [])
        r2 = copy.deepcopy(r); r2['direct_url'] = 'https://example.test/job?id=7&bu=2&source=tracker'
        r2['role_key'] = identity(r2)
        self.assertTrue(any('duplicate official role' in e for e in self.check(roles=[r, r2])))

    def test_malformed_contract_returns_errors(self):
        for p in [None, [], {}, {'hard_scope': None, 'axes': None}]:
            self.assertTrue(validate_snapshot(p, [None, {}], [None], None))

    def test_prepared_stage_cannot_be_login_only(self):
        r = role(); r['stage'] = 'prepared'
        self.assertTrue(any('completed fill evidence' in e for e in self.check(roles=[r])))

    def test_discovered_pending_clue_cannot_bypass_form_gates(self):
        r = role(); r.update(stage='discovered', disposition='pending')
        self.assertEqual(self.check(roles=[r]), [])
        r['execution']['fill'] = {'status': 'in_progress', 'evidence_ref': 'synthetic'}
        self.assertTrue(any('cannot perform form' in e for e in self.check(roles=[r])))

    def test_source_truth_is_not_inferred_from_jd_words(self):
        r = role(); r.update(jd_full_text='Synthetic externally unverified text', jd_basis='manual declared evidence')
        self.assertEqual(self.check(roles=[r]), [])  # Structural validity is not source acceptance.


class IntegrationTests(unittest.TestCase):
    def test_assignment_contract_valid_and_rejects_old_delta_or_missing_report(self):
        sub = 'application-parallel-execution' if (ROOT/'scripts/validate_batch.py').exists() else 'campus-job-parallel-search'
        validator = load(ROOT/'skills'/sub/'scripts/validate_assignments.py')
        workflow = validator.WORKFLOW
        a = dict(assignment_id='a', company_key='synthetic-ats', owner='one', output_dir='evidence/one/a',
                 preference_version='p2', recruiting_entity='synthetic-entity', operation_scope='synthetic-account', source_scope='synthetic-plan')
        r = role(); r.update(assignment_id='a', owner='one')
        m = dict(workflow=workflow, batch_id='synthetic', query_version='q2', mode='single',
                 mode_approval_ref='synthetic-mode', execution_channel='manual_handoff', participants=['one'],
                 coordinator='one', output_roots={'one':'evidence/one'}, source_hash='a'*64,
                 assignment_version=1, assignments=[a], roles=[r], preferences=preferences(), prior_roles=[])
        self.assertEqual(validator.validate(m), [])
        delta = dict(batch_id='synthetic', query_version='q2', assignment_version=1, source_hash='a'*64,
                     assignment_id='a', owner='one', preference_version='p1')
        m['deltas'] = [delta]
        self.assertTrue(any('stale delta preference_version' in e for e in validator.validate(m)))
        m['deltas'] = []; r.pop('pre_form_report_ref')
        self.assertTrue(any('pre_form_report_ref' in e for e in validator.validate(m)))
        r['pre_form_report_ref'] = 'synthetic-report'; a['preference_version'] = 'p1'
        self.assertTrue(any('stale assignment' in e for e in validator.validate(m)))

    def test_research_progress_preserves_coverage_and_checks_new_contract(self):
        path = ROOT/'scripts/validate_research_progress.py'
        if not path.exists():
            self.skipTest('research validator belongs to campus package')
        validator = load(path)
        c = company(); c.update(id='synthetic-ats', raw_status='尚未推进', result_status='尚未推进',
             coverage_bucket='first-pass-visited-but-unresolved', checked_at='synthetic-date',
             cycle_ref='synthetic-cycle', attempt_refs=['synthetic-attempt'], coverage_refs=[],
             remaining_gaps=['one JD unread'], merge_status='not_merged', preference_version='p2')
        m = dict(records=[c], roles=[role()], preferences=preferences(), prior_roles=[])
        self.assertEqual(validator.validate(m), [])
        c['coverage_bucket'] = 'first-pass-unvisited'
        self.assertTrue(any('unvisited contradicts' in e for e in validator.validate(m)))
        c['coverage_bucket'] = 'first-pass-covered'
        self.assertTrue(any('covered requires' in e for e in validator.validate(m)))
        c['coverage_bucket'] = 'first-pass-visited-but-unresolved'; c['preference_version'] = 'p1'
        self.assertTrue(any('stale preference_version' in e for e in validator.validate(m)))


if __name__ == '__main__':
    unittest.main()
