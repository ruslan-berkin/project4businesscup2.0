"""Gemini JSON transport and opt-in cross-provider fallback; no live calls."""
import json
import os
import unittest
from unittest.mock import Mock, patch
import requests
import app


class GeminiTests(unittest.TestCase):
    def setUp(self):
        app._AI_COOLDOWNS.clear()
        self.client = app.app.test_client()
        env = patch.dict(os.environ, {'AI_ENABLED':'true','AI_PROVIDER':'gemini',
            'AI_FALLBACK_PROVIDERS':'','GEMINI_API_KEY':'test-gemini','GEMINI_API_KEYS':'',
            'GROQ_API_KEY':'test-groq','GROQ_API_KEYS':'','GEMINI_MODEL':'gemini-3.5-flash-lite'})
        env.start(); self.addCleanup(env.stop)

    def answer(self):
        data=dict(narrative='Lena keeps the letter.',options=['Read','Help Lena','Rest'],delta=[0,-3,1],
                  memory='Lena trusts the traveller.',item=None,ending=False)
        return Mock(json=lambda:{'candidates':[{'finishReason':'STOP','content':{'parts':[{'text':json.dumps(data)}]}}]})

    @patch('app.requests.post')
    def test_gemini_start_and_turn(self, call):
        call.return_value=self.answer()
        state=self.client.post('/api/start',json={'mode':'endless','genre':'detective'}).get_json()
        self.assertEqual(state['engine'],'gemini')
        result=self.client.post('/api/action',json={'story_token':state['story_token'],'choice_idx':1})
        self.assertEqual(result.status_code,200)
        self.assertEqual(result.get_json()['stats']['KPI'],97)
        args=call.call_args
        self.assertIn(':generateContent',args.args[0]); self.assertNotIn('test-gemini',args.args[0])
        self.assertEqual(args.kwargs['headers']['x-goog-api-key'],'test-gemini')
        body=args.kwargs['json']
        self.assertEqual(body['generationConfig']['responseJsonSchema'],app.AI_SCHEMA)
        context=json.loads(body['contents'][0]['parts'][0]['text'])
        self.assertEqual(context['action'],'Help Lena');self.assertEqual(context['memory'],'Lena trusts the traveller.')
        self.assertEqual(context['difficulty']['level'],0)

    @patch('app.requests.post')
    def test_groq_limit_falls_back_only_when_enabled(self, call):
        response=requests.Response();response.status_code=429;response.headers['Retry-After']='60'
        error=requests.HTTPError('private body',response=response)
        with patch.dict(os.environ,{'AI_PROVIDER':'groq'}):
            call.side_effect=error
            self.assertEqual(self.client.post('/api/start',json={'mode':'premium'}).status_code,502)
            self.assertEqual(call.call_count,1)
            app._AI_COOLDOWNS.clear();call.reset_mock()
            call.side_effect=[error,self.answer()]
            with patch.dict(os.environ,{'AI_FALLBACK_PROVIDERS':'gemini'}):
                state=self.client.post('/api/start',json={'mode':'endless','genre':'detective'}).get_json()
                self.assertEqual(state['engine'],'gemini')
                self.assertEqual(call.call_count,2)
                call.side_effect=None;call.return_value=self.answer()
                result=self.client.post('/api/action',json={'story_token':state['story_token'],'choice_idx':0})
                self.assertEqual(result.status_code,200)
                self.assertEqual(result.get_json()['story_turn'],1)

    @patch('app.requests.post')
    def test_invalid_gemini_response_never_applies_a_turn(self, call):
        call.return_value=self.answer()
        state=self.client.post('/api/start',json={'mode':'endless'}).get_json()
        for payload in ({}, {'candidates':[{'finishReason':'MAX_TOKENS'}]}, {'candidates':[{'finishReason':'STOP','content':{'parts':[{'text':'{}'}]}}]}):
            call.return_value=Mock(json=lambda:payload)
            result=self.client.post('/api/action',json={'story_token':state['story_token'],'choice_idx':0})
            self.assertEqual(result.status_code,502)
            self.assertEqual(app.story_signer().loads(state['story_token'])['turn'],0)

if __name__=='__main__': unittest.main()
