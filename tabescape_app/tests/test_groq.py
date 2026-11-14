"""Groq transport and story continuity: every outbound request is mocked."""
import json
import os
import unittest
from unittest.mock import Mock, patch

import requests
import app


class GroqContractTests(unittest.TestCase):
    def setUp(self):
        self.client = app.app.test_client()
        env = patch.dict(os.environ, {
            'AI_PROVIDER': 'groq', 'AI_ENABLED': 'true',
            'GROQ_API_KEY': 'test-only-groq', 'GROQ_MODEL': 'openai/gpt-oss-20b',
            'OPENAI_API_KEY': 'test-only-openai',
        })
        env.start()
        self.addCleanup(env.stop)
        outbound = patch('app.requests.post')
        self.call = outbound.start()
        self.addCleanup(outbound.stop)
        self.call.return_value = self.answer()

    def answer(self, **changes):
        data = dict(narrative='Lena opens the tower.', options=['Read', 'Follow Lena', 'Rest'],
                    delta=[-2, -3, 1], memory='Lena has the ledger.', item=None, ending=False)
        data.update(changes)
        return Mock(json=lambda: {'choices': [{'finish_reason': 'stop', 'message': {'content': json.dumps(data)}}]})

    def start(self, mode='endless'):
        response = self.client.post('/api/start', json={'mode': mode, 'genre': 'detective', 'language': 'en'})
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()

    def turn(self, state, **changes):
        return self.client.post('/api/action', json={'story_token': state['story_token'], 'choice_idx': 1, **changes})

    def test_endless_continuity_transport_and_inventory(self):
        self.call.return_value = self.answer(ending=True, item='Kit')
        state = self.start()
        self.assertEqual(state['engine'], 'groq')
        self.assertEqual(state['stats']['ROI'], 100)
        for _ in range(5):
            response = self.turn(state)
            self.assertEqual(response.status_code, 200)
            state = response.get_json()
            self.assertFalse(state['game_over'])
        self.assertEqual(state['stats']['ROI'], 90)
        request = self.call.call_args.kwargs['json']
        self.assertEqual(request['model'], 'openai/gpt-oss-20b')
        self.assertTrue(request['response_format']['json_schema']['strict'])
        context = json.loads(request['messages'][1]['content'])
        self.assertEqual(context['memory'], 'Lena has the ledger.')
        self.assertEqual(context['action'], 'Follow Lena')
        self.assertEqual(len(context['history']), 4)
        self.assertIn('Kit', context['inventory'])
        for args in self.call.call_args_list:
            self.assertEqual(args.args[0], 'https://api.groq.com/openai/v1/chat/completions')
            self.assertEqual(args.kwargs['headers']['Authorization'], 'Bearer test-only-groq')
        calls = self.call.call_count
        healed = self.turn(state, choice_idx=-1, action='Use item: Kit')
        self.assertEqual(healed.status_code, 200)
        self.assertGreater(healed.get_json()['stats']['ROI'], 90)
        self.assertEqual(self.call.call_count, calls)

    def test_finite_story_has_finale(self):
        self.call.return_value = self.answer(ending=True)
        state = self.start('premium')
        for i in range(4):
            state = self.turn(state).get_json()
            self.assertEqual(state['game_over'], i == 3)
        self.assertTrue(state['won'])
        self.assertEqual(state['options'], [])

    def test_missing_key_disabled_and_unknown_provider_use_demo(self):
        for settings in ({'GROQ_API_KEY': ''}, {'GROQ_API_KEY': '  '}, {'AI_ENABLED': 'false'}, {'AI_PROVIDER': 'unknown'}):
            with self.subTest(settings=list(settings)), patch.dict(os.environ, settings):
                catalog = self.client.get('/api/catalog').get_json()
                self.assertFalse(catalog['premium_ai'])
                self.assertIsNone(catalog['ai_provider'])
                self.assertEqual(self.start()['engine'], 'demo')
        self.call.assert_not_called()

    def test_free_stories_do_not_call_groq_and_catalog_contains_no_key(self):
        state = self.start('free')
        self.assertEqual(self.turn(state).status_code, 200)
        self.call.assert_not_called()
        catalog = self.client.get('/api/catalog')
        self.assertEqual(catalog.get_json()['ai_provider'], 'groq')
        self.assertTrue(catalog.get_json()['premium_ai'])
        self.assertNotIn('test-only', catalog.get_data(as_text=True))

    def test_http_failures_and_timeout_leave_token_retryable(self):
        state = self.start()
        for error in (requests.Timeout('test'), requests.HTTPError('401'), requests.HTTPError('429')):
            with self.subTest(error=str(error)):
                self.call.side_effect = error
                failed = self.turn(state)
                self.assertEqual(failed.status_code, 502)
                self.assertNotIn('story_token', failed.get_json())
                self.assertEqual(app.story_signer().loads(state['story_token'])['turn'], 0)
        self.call.side_effect = None
        retried = self.turn(state).get_json()
        self.assertEqual(retried['story_turn'], 1)
        self.assertEqual(retried['stats']['ROI'], 98)

    def test_invalid_or_incomplete_responses_are_not_applied(self):
        invalid_payloads = [None, [], {}, {'choices': []}, {'choices': [None]},
                            {'choices': [{'finish_reason': 'length', 'message': {'content': '{}'}}]},
                            {'choices': [{'finish_reason': 'stop', 'message': {'refusal': 'No'}}]}]
        for content in ('not json', '[]', '{}', 'null'):
            invalid_payloads.append({'choices': [{'finish_reason': 'stop', 'message': {'content': content}}]})
        for payload in invalid_payloads:
            with self.subTest(payload=payload):
                self.call.return_value = Mock(json=lambda: payload)
                self.assertEqual(self.client.post('/api/start', json={'mode': 'premium'}).status_code, 502)
        for changes in ({'options': ['One']}, {'delta': [True, 0, 1]}, {'item': {}}, {'ending': 'true'}):
            self.call.return_value = self.answer(**changes)
            self.assertEqual(self.client.post('/api/start', json={'mode': 'premium'}).status_code, 502)

    def test_provider_switch_does_not_silently_move_existing_story(self):
        state = self.start()
        self.call.reset_mock()
        with patch.dict(os.environ, {'AI_PROVIDER': 'openai'}):
            self.assertEqual(self.turn(state).status_code, 502)
        self.call.assert_not_called()

    def test_legacy_openai_enabled_setting_remains_compatible(self):
        with patch.dict(os.environ, {'AI_PROVIDER': 'openai', 'OPENAI_ENABLED': 'true'}):
            os.environ.pop('AI_ENABLED', None)
            self.assertTrue(app.ai_ready())
            os.environ['OPENAI_ENABLED'] = 'false'
            self.assertFalse(app.ai_ready())


if __name__ == '__main__':
    unittest.main()
