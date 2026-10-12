"""Public CLI defaults must never overwrite V2 with historical SB protections."""
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import run_categorization as pipeline

class CurrentEntrypointTests(unittest.TestCase):
    def test_default_final_delivery_routes_to_current_model_and_visualization(self):
        with patch.object(sys,'argv',['run_categorization.py']),patch('subprocess.run') as invoke:
            pipeline.main()
        self.assertEqual(len(invoke.call_args_list),2)
        self.assertTrue(invoke.call_args_list[0].args[0][1].endswith('national_categorization.py'))
        self.assertTrue(invoke.call_args_list[1].args[0][1].endswith('prepare_data.py'))
        self.assertTrue(all(call.kwargs['check'] for call in invoke.call_args_list))

    def test_category_phase_also_uses_current_model_without_old_sb_promotions(self):
        with patch.object(sys,'argv',['run_categorization.py','--through','F7']),patch('subprocess.run') as invoke:
            pipeline.main()
        self.assertEqual(invoke.call_count,1)
        self.assertTrue(invoke.call_args.args[0][1].endswith('national_categorization.py'))

if __name__=='__main__':unittest.main()
