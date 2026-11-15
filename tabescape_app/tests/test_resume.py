"""Saved games resume without new turns, new items or AI requests."""
import os
import unittest
from unittest.mock import patch
import app


@patch.dict(os.environ, {'AI_ENABLED': 'false', 'OPENAI_ENABLED': 'false'})
class ResumeTests(unittest.TestCase):
    def setUp(self):
        self.client = app.app.test_client()

    def test_all_modes_restore_exact_state_and_can_continue(self):
        for mode in ('free', 'premium', 'endless'):
            with self.subTest(mode=mode), patch('app.requests.post') as ai:
                state = self.client.post('/api/start', json={'mode': mode, 'campaign_id': 'oakhaven'}).get_json()
                for choice in (0, 0, 1, 0):
                    state = self.client.post('/api/action', json={'story_token': state['story_token'], 'choice_idx': choice}).get_json()
                token = state['story_token']
                for _ in range(2):
                    response = self.client.post('/api/resume', json={'story_token': token})
                    self.assertEqual(response.status_code, 200)
                    restored = response.get_json()
                    for field in ('narrative', 'stats', 'options', 'story_turn', 'scene_id', 'mode', 'game_over'):
                        self.assertEqual(restored[field], state[field])
                    self.assertIsNone(restored['inventory_add'])
                    self.assertEqual(restored['inventory'], app.story_signer().loads(token)['inventory'])
                next_turn = self.client.post('/api/action', json={'story_token': restored['story_token'], 'choice_idx': 0}).get_json()
                self.assertEqual(next_turn['story_turn'], state['story_turn'] + 1)
                ai.assert_not_called()

    def test_ai_memory_and_finished_state_survive_without_generation(self):
        state = self.client.post('/api/start', json={'mode': 'free'}).get_json()
        saved = app.story_signer().loads(state['story_token'])
        saved.update(engine='groq', mode='endless', memory='Lena remembers the promise', over=True, won=False, options=[])
        with patch('app.generate_ai') as generate:
            response = self.client.post('/api/resume', json={'story_token': app.story_signer().dumps(saved)})
        self.assertEqual(response.status_code, 200)
        restored = response.get_json()
        self.assertTrue(restored['game_over'])
        self.assertEqual(app.story_signer().loads(restored['story_token']), saved)
        generate.assert_not_called()

    def test_invalid_and_old_secret_are_recoverable_errors(self):
        state = self.client.post('/api/start', json={'mode': 'free'}).get_json()
        for payload in (None, [], {}, {'story_token':'damaged'}, {'story_token':app.story_signer().dumps({})}):
            self.assertEqual(self.client.post('/api/resume', json=payload).status_code, 400)
        with patch.object(app, 'STORY_SECRET', 'different-server-secret'):
            self.assertEqual(self.client.post('/api/resume', json={'story_token':state['story_token']}).status_code, 400)
