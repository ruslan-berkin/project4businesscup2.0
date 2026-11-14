"""Branches, survival and the server-side OpenAI contract (no real API requests)."""
import json
import os
import unittest
from unittest.mock import Mock, patch
import requests
import app


@patch.dict(os.environ, {'OPENAI_ENABLED': 'false', 'OPENAI_API_KEY': ''})
class CampaignTests(unittest.TestCase):
    def setUp(self):
        self.client = app.app.test_client()

    def start(self, **kwargs):
        response = self.client.post('/api/start', json={'campaign_id': 'oakhaven', 'mode': 'free', 'language': 'en', **kwargs})
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()

    def turn(self, state, choice=0, **kwargs):
        response = self.client.post('/api/action', json={'story_token': state['story_token'], 'choice_idx': choice, **kwargs})
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()

    def test_eight_campaigns_in_two_languages_have_three_endings(self):
        campaigns = self.client.get('/api/catalog').get_json()['campaigns']
        self.assertEqual(len(campaigns), 8)
        starts = set()
        for campaign in campaigns:
            for language in ('ru', 'en'):
                endings = []
                for final_choice in (0, 1, 2):
                    with self.subTest(campaign=campaign['id'], language=language, ending=final_choice):
                        state = self.start(campaign_id=campaign['id'], language=language)
                        starts.add(state['narrative'])
                        for choice in (0, 0, 0, final_choice):
                            state = self.turn(state, choice)
                        self.assertTrue(state['game_over'])
                        self.assertEqual(state['story_turn'], 4)
                        self.assertEqual(state['won'], final_choice == 0)
                        self.assertEqual(state['options'], [])
                        endings.append(state['narrative'])
                self.assertEqual(len(set(endings)), 3)
        self.assertEqual(len(starts), 16)

    def test_evidence_affects_ending_and_risky_actions_can_kill(self):
        state = self.start()
        for choice in (2, 1, 1, 0):
            state = self.turn(state, choice)
        self.assertTrue(state['game_over'])
        self.assertFalse(state['won'])
        dead = self.start()
        for _ in range(4):
            if not dead['game_over']:
                dead = self.turn(dead, -1, action='Attack the guard')
        self.assertTrue(dead['game_over'])
        self.assertEqual(dead['stats']['ROI'], 0)
        self.assertEqual(dead['story_turn'], self.turn(dead)['story_turn'])

    def test_inventory_and_typed_action(self):
        state = self.start()
        for _ in range(3):
            state = self.turn(state)
        item = state['inventory_add']
        self.assertTrue(item)
        restored = self.turn(state, -1, action='Use item: ' + item)
        self.assertGreater(restored['stats']['KPI'], state['stats']['KPI'])
        self.assertEqual(restored['scene_id'], state['scene_id'])
        response = self.client.post('/api/action', json={'story_token': restored['story_token'], 'action': 'Use item: ' + item})
        self.assertEqual(response.status_code, 400)
        custom = self.turn(self.start(), -1, action='Inspect the canal')
        self.assertEqual(custom['story_turn'], 1)
        self.assertNotEqual(custom['scene_id'], 'start')

    def test_endless_passes_eighteen_oclock_then_ends_only_on_death(self):
        state = self.start(mode='endless', genre='horror')
        self.assertEqual(state['engine'], 'demo')
        self.assertEqual(state['campaign_id'], 'room6')
        for i in range(22):
            state = self.turn(state, 0 if i % 4 in (0, 3) else 1)
            self.assertFalse(state['game_over'], i)
        self.assertGreater(state['story_turn'], 18)
        for _ in range(20):
            if state['game_over']:
                break
            state = self.turn(state, -1, action='Attack the shadow')
        self.assertTrue(state['game_over'])
        self.assertEqual(state['stats']['ROI'], 0)
        self.assertFalse(state['won'])

    def test_untrusted_stats_invalid_token_and_unknown_mode(self):
        state = self.start()
        next_state = self.turn(state, stats={'ROI': 0, 'KPI': 0, 'Anxiety': 100})
        self.assertFalse(next_state['game_over'])
        response = self.client.post('/api/action', json={'story_token': state['story_token']+'tampered'})
        self.assertEqual(response.status_code, 400)
        for change in ({'mode':'unknown'}, {'campaign_id':'missing'}, {'genre':'missing'}):
            response = self.client.post('/api/start', json={'campaign_id':'oakhaven', **change})
            self.assertEqual(response.status_code, 400)

    @patch('app.requests.post')
    def test_free_never_calls_openai(self, call):
        with patch.dict(os.environ, {'OPENAI_ENABLED':'true', 'OPENAI_API_KEY':'test-only'}):
            self.turn(self.start())
        call.assert_not_called()


class OpenAIContractTests(unittest.TestCase):
    def setUp(self):
        self.client = app.app.test_client()
        self.env = patch.dict(os.environ, {'OPENAI_ENABLED':'true', 'OPENAI_API_KEY':'test-only', 'OPENAI_MODEL':'gpt-5-mini'})
        self.env.start()
        self.addCleanup(self.env.stop)

    def answer(self, **changes):
        data = {'narrative':'The archivist keeps your promise.', 'options':['Read the letter','Follow Lena','Rest'],
                'delta':[-2,-3,1], 'memory':'Lena trusts the hero; the missing ledger is still in the tower.', 'item':None, 'ending':False, **changes}
        return Mock(json=lambda: {'status':'completed','output':[{'content':[{'type':'output_text','text':json.dumps(data)}]}]})

    @patch('app.requests.post')
    def test_openai_context_continues_endless_ignores_model_finale(self, call):
        call.return_value = self.answer(ending=True)
        state = self.client.post('/api/start', json={'mode':'endless','genre':'detective','language':'en','role':'Ana'}).get_json()
        self.assertEqual(state['engine'], 'openai')
        self.assertEqual(state['stats']['ROI'], 100)
        for _ in range(5):
            response = self.client.post('/api/action', json={'story_token':state['story_token'], 'choice_idx':1})
            self.assertEqual(response.status_code, 200)
            state = response.get_json()
            self.assertFalse(state['game_over'])
        args = call.call_args.kwargs
        self.assertEqual(call.call_args.args[0], 'https://api.openai.com/v1/responses')
        self.assertFalse(args['json']['store'])
        self.assertEqual(args['json']['text']['format']['type'], 'json_schema')
        context = json.loads(args['json']['input'])
        self.assertIn('Lena trusts', context['memory'])
        self.assertTrue(context['history'])
        self.assertEqual(context['action'], 'Follow Lena')
        self.assertEqual(context['role'], 'Ana')
        self.assertEqual(state['stats']['ROI'], 90)

    @patch('app.requests.post')
    def test_ai_finale_and_failures_do_not_consume_turn(self, call):
        call.return_value = self.answer(item='Emergency kit')
        state = self.client.post('/api/start', json={'mode':'premium', 'genre':'scifi'}).get_json()
        self.assertIn('Emergency kit', app.story_signer().loads(state['story_token'])['inventory'])
        call.side_effect = requests.Timeout('test')
        failed = self.client.post('/api/action', json={'story_token':state['story_token'], 'choice_idx':0})
        self.assertEqual(failed.status_code, 502)
        call.side_effect = None
        call.return_value = self.answer(ending=True)
        for i in range(4):
            state = self.client.post('/api/action', json={'story_token':state['story_token'], 'choice_idx':0}).get_json()
            self.assertEqual(state['story_turn'], i+1)
        self.assertTrue(state['game_over'])
        self.assertTrue(state['won'])
        call.return_value = self.answer(options=['Only one'])
        failed = self.client.post('/api/start', json={'mode':'premium'})
        self.assertEqual(failed.status_code, 502)

if __name__ == '__main__':
    unittest.main()
