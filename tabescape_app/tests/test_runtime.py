import base64
import importlib.util
import io
import os
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch
import zipfile
import zlib

import app


class RuntimeTests(unittest.TestCase):
    def test_embedded_assets_match_and_restore(self):
        root = Path(app.BASE_DIR)
        for path, data in app._EMBEDDED_FILES.items():
            self.assertEqual((root / path).read_bytes(), zlib.decompress(base64.b64decode(data)), path)
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'app.py'
            shutil.copy(root / 'app.py', target)
            spec = importlib.util.spec_from_file_location('standalone_tabescape', target)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            client = module.app.test_client()
            self.assertEqual(client.get('/').status_code, 200)
            with client.get('/static/js/app.js') as response:
                self.assertIn(b'initWorkspace', response.data)
            for resource in ['/static/guides/excel.jpeg', '/api/workbook/default',
                             '/static/js/tour.js', '/static/images/mira.webp',
                             '/static/images/workspace-background-1.webp', '/api/catalog']:
                with client.get(resource) as response:
                    self.assertEqual(response.status_code, 200)
            self.assertIn('You wake', client.post('/api/start', json={'language': 'en'}).get_json()['narrative'])
            self.assertEqual(len(client.get('/api/catalog').get_json()['campaigns']), 8)
            self.assertEqual(client.post('/api/start', json={'campaign_id': 'row17'}).status_code, 200)

    @patch.dict(os.environ, {}, clear=True)
    def test_procedural_start_and_action(self):
        client = app.app.test_client()
        start = client.post('/api/start', json={}).get_json()
        self.assertEqual(start['stats'], {'ROI': 100, 'KPI': 100, 'Anxiety': 10})
        self.assertEqual(len(start['options']), 3)
        action = client.post('/api/action', json={
            'action': start['options'][0], 'choice_idx': 0,
            'scene_id': start['scene_id'], 'stats': start['stats']
        }).get_json()
        self.assertNotEqual(action['stats'], start['stats'])
        self.assertNotEqual(action['narrative'], start['narrative'])

    def test_uploaded_interface_uses_server_and_local_assets(self):
        client = app.app.test_client()
        html = client.get('/').get_data(as_text=True)
        for name in ('assets', 'office-i18n', 'tour', 'allowance', 'app'):
            self.assertIn('/static/js/' + name + '.js', html)
            with client.get('/static/js/' + name + '.js') as response:
                self.assertEqual(response.status_code, 200)
        with client.get('/static/js/app.js') as response:
            js = response.get_data(as_text=True)
        self.assertNotIn('localAdventure', js)
        self.assertIn('await fetch(url', js)
        self.assertIn('/api/workbook/default', js)
        self.assertIn('/api/export_xlsx', js)
        for name in ('mira.webp', 'workspace-background-1.webp'):
            with client.get('/static/images/' + name) as response:
                self.assertEqual(response.status_code, 200)

    @patch.dict(os.environ, {}, clear=True)
    def test_english_uses_same_procedural_rules(self):
        client = app.app.test_client()
        start = client.post('/api/start', json={'language': 'en'}).get_json()
        self.assertIn('You wake', start['narrative'])
        action = client.post('/api/action', json={
            'language': 'en', 'action': start['options'][0], 'choice_idx': 0,
            'scene_id': start['scene_id'], 'stats': start['stats']
        }).get_json()
        self.assertEqual(action['stats'], {'ROI': 55, 'KPI': 80, 'Anxiety': 25})
        self.assertIn('ravine', action['narrative'])
        attack = client.post('/api/action', json={
            'language': 'en', 'action': 'Attack the dragon', 'choice_idx': -1,
            'scene_id': action['scene_id'], 'stats': action['stats']
        }).get_json()
        self.assertEqual(attack['stats']['ROI'], 15)
        item = client.post('/api/action', json={
            'language': 'en', 'action': 'Use item: Potion', 'choice_idx': -1,
            'scene_id': attack['scene_id'], 'stats': attack['stats']
        }).get_json()
        self.assertEqual(item['stats']['ROI'], 40)

    def test_default_and_export(self):
        client = app.app.test_client()
        source = client.get('/api/workbook/default')
        self.assertEqual(source.status_code, 200)
        response = client.post('/api/export_xlsx', json={
            'mode': 'imported', 'filename': 'check.xlsx',
            'source_b64': base64.b64encode(source.data).decode(),
            'edits': {'imp0': {'2:2': 150000}}, 'styles': {}
        })
        source.close()
        self.assertEqual(response.status_code, 200)
        with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
            sheet = archive.read('xl/worksheets/sheet1.xml')
            self.assertIn(b'150000', sheet)
            self.assertIn(b'C2-B2', sheet)


if __name__ == '__main__':
    unittest.main()
