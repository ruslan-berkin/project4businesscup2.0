"""Meaningful branches, defeat aftermath, escalating survival and reserve transport."""
import json
import os
import unittest
from unittest.mock import Mock, patch
import requests
import app


class StoryDepthTests(unittest.TestCase):
    def setUp(self):
        self.client = app.app.test_client()
        app._AI_COOLDOWNS.clear()

    def play(self, campaign, choices, lang='ru'):
        state = self.client.post('/api/start', json={'campaign_id': campaign, 'mode': 'free', 'language': lang}).get_json()
        for choice in choices:
            state = self.client.post('/api/action', json={'story_token': state['story_token'], 'choice_idx': choice}).get_json()
        return state

    def test_clues_without_trust_are_not_enough_and_every_story_has_aftermath(self):
        for cid, campaign in app.campaign_data().items():
            for lang in ('ru', 'en'):
                with self.subTest(campaign=cid, lang=lang):
                    self.assertFalse(self.play(cid, [0]*6, lang)['won'])
                    self.assertTrue(self.play(cid, [0,0,1,0,0,0], lang)['won'])
                    dead = self.play(cid, [2]*6, lang)
                    self.assertTrue(dead['game_over'])
                    self.assertIn(campaign['endings']['loss'][lang], dead['narrative'])
                    self.assertEqual(dead['options'], [])

    @patch('app.generate_ai')
    def test_pressure_is_gradual_and_recovery_can_extend_survival(self, generate):
        def simulate(delta):
            generate.return_value = dict(narrative='Lena remembers your promise.', options=['Listen','Help','Rest'],
                memory='Lena is your ally.', item=None, ending=False, delta=delta)
            state = dict(mode='endless', engine='groq', stats={'ROI':100,'KPI':100,'Anxiety':10},
                         turn=0, over=False, language='en')
            samples = {}
            for turn in range(1,36):
                state['turn'] = turn
                app.ai_story_turn(state, 'Rest' if delta[1]>0 else 'Walk onward')
                samples[turn] = dict(state['stats'])
                if state['over']: break
            return state, samples
        dead, samples = simulate([0,-3,1])
        self.assertEqual(samples[9]['KPI'], 73)
        self.assertEqual(samples[10]['KPI'], 68)
        self.assertTrue(25 <= dead['turn'] <= 30)
        self.assertTrue(dead['over'])
        self.assertIn('world continues', dead['narrative'])
        survivor, _ = simulate([5,20,-10])
        self.assertFalse(survivor['over'])
        self.assertEqual(survivor['turn'], 35)

    @patch.dict(os.environ, {'GROQ_API_KEY':'primary-test', 'GROQ_API_KEYS':'reserve-test, third-test, fourth-test'})
    @patch('app.requests.post')
    def test_429_switches_to_reserve_and_respects_cooldown(self, call):
        response = requests.Response(); response.status_code = 429; response.headers['Retry-After'] = '90'
        error = requests.HTTPError('private body', response=response)
        success = Mock()
        call.side_effect = [error, success]
        self.assertIs(app.ai_post('groq', 'https://api.groq.com/openai/v1/chat/completions', {}), success)
        self.assertEqual([c.kwargs['headers']['Authorization'] for c in call.call_args_list],
                         ['Bearer primary-test','Bearer reserve-test'])
        call.reset_mock(); call.side_effect = None; call.return_value = success
        app.ai_post('groq','https://api.groq.com/openai/v1/chat/completions',{})
        self.assertEqual(call.call_args.kwargs['headers']['Authorization'],'Bearer reserve-test')
        app._AI_COOLDOWNS.clear(); call.reset_mock(); call.side_effect = error
        with self.assertRaises(requests.HTTPError): app.ai_post('groq','unused',{})
        self.assertEqual(call.call_count, 3, 'bound attempts even with many reserve keys')

    @patch.dict(os.environ, {'GROQ_API_KEY':'primary-test', 'GROQ_API_KEYS':'reserve-test'})
    @patch('app.requests.post')
    def test_all_cooling_down_makes_no_requests_and_timeout_is_not_replayed(self, call):
        call.side_effect = requests.Timeout('test')
        with self.assertRaises(requests.Timeout): app.ai_post('groq','unused',{})
        self.assertEqual(call.call_count, 1)
        import hashlib, time
        for key in app.ai_keys('groq'):
            app._AI_COOLDOWNS[hashlib.sha256(('groq'+key).encode()).hexdigest()] = time.monotonic()+60
        call.reset_mock()
        with self.assertRaises(requests.HTTPError) as caught: app.ai_post('groq','unused',{})
        self.assertEqual(caught.exception.response.status_code, 429)
        call.assert_not_called()

if __name__ == '__main__': unittest.main()
