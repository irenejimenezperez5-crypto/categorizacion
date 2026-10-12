"""Accounting and edge-case tests; run with unittest discover."""
import sys
from pathlib import Path
import unittest
import numpy as np
import pandas as pd
from unittest.mock import patch
from scipy import sparse

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from categorization_core import code, allocate_cents, age_demand, distance_shares, haversine_matrix, categories


class AttributionTests(unittest.TestCase):
    def test_missing_iqvia_ids_do_not_collapse(self):
        for value in [None, np.nan, "NULL", "#N/A", "0000000", "not assigned"]:
            self.assertIsNone(code(value, 7))
        self.assertEqual(code(12345., 7), "0012345")

    def test_conserve_cents_with_unequal_weights_and_ties(self):
        self.assertEqual(allocate_cents(10001, np.array([1., 1., 1.])).tolist(), [3334, 3334, 3333])
        rng = np.random.default_rng(2027)
        for _ in range(30):
            values = allocate_cents(987654321, rng.random(63))
            self.assertEqual(int(values.sum()), 987654321)
            self.assertTrue((values >= 0).all())

    def test_exclusive_age_groups(self):
        self.assertEqual(age_demand(np.array([[10., 20., 30., 40., 50.]]))[0], 320.)

    def test_uncertain_store_is_one_competitor_and_zero_distance(self):
        d = np.array([[0., 0., 0., 0.]])
        share = distance_shares(d, np.array([0, 1]), np.array([1., 1/3, 1/3, 1/3]), np.array([1.]))
        np.testing.assert_allclose(share, [.5, .5])
        np.testing.assert_allclose(haversine_matrix(np.array([[-3., 40.]]), np.array([[-3., 40.]])), [[0.]])

    def test_gradual_border_capture_and_distance_order(self):
        share = distance_shares(np.array([[50., 60., 1000.]]), np.arange(3), np.ones(3), np.ones(1))
        self.assertGreater(share[1], .4)  # competitor across an arbitrary section/SB border
        self.assertGreater(share[0], share[1])
        self.assertGreater(share[1], share[2])
        self.assertAlmostEqual(share.sum(), 1.)

    def test_cumulative_euros_prior_and_high_sb_rules(self):
        stores = pd.DataFrame({"establishment_id": list("abcdef"), "ccaa_modelo": ["X"]*6,
                               "sb_modelo": ["high"]*2 + ["other"]*4,
                               "potential_cents": [51, 20, 10, 9, 7, 3]})
        result = categories(stores, pd.Series({"high": 10000, "other": 100}))
        self.assertEqual(result.categoria_base.tolist(), ["A", "C", "C", "D", "D", "D"])
        self.assertEqual(result.potential_cents.sum(), 100)
        lone = stores.iloc[:1].copy()
        result = categories(lone, pd.Series({"high": 10000, "other": 100}))
        self.assertEqual(result.categoria_final.iloc[0], "B")
        stores.loc[5, "sb_modelo"] = "high"
        result = categories(stores, pd.Series({"high": 10000, "other": 100}))
        self.assertEqual(result.categoria_final.iloc[5], "C")

    def test_rural_municipality_without_stores_uses_origin_sb_neutrally(self):
        import run_categorization as run
        sections = pd.DataFrame({"sb":["SB01","SB02"],"municipio":["SIN_PUNTOS","CON_PUNTO"],"degurba":[3,3],"pop_total":[100.,200.]})
        from categorization_core import AGE_COLUMNS
        for c in AGE_COLUMNS: sections[c]=[10.,20.]
        stores = pd.DataFrame({"sb_modelo":["SB01","SB01","SB02"]})
        market = {"budgets":pd.Series([101,10000],index=["SB01","SB02"])}
        geo = {"sections":sections,"neighbors":{},"points":[]}
        loc = {"stores":stores,"placement_xy":np.array([[0.,0.],[1.,1.],[2.,2.]]),"offsets":np.arange(4),"municipal_presence":{"CON_PUNTO":{2:1.}}}
        artifacts={}
        with patch.object(run,"load",side_effect=lambda name:{"geodemography":geo,"locations":loc,"market":market}[name]),patch.object(run,"dump",side_effect=lambda name,obj:artifacts.update({name:obj})),patch.object(run,"write_json"),patch.object(run,"log"),patch.object(run.sparse,"save_npz"):
            report=run.f5({"distance_exponent":2.,"age_weights":[1.,2.,3.]})
        self.assertEqual(report["rural_no_municipal_store_sections"],1)
        observed=artifacts["attribution"]["flow_cents"].toarray()
        # An odd-cent neutral budget must award the leftover cent to one store.
        np.testing.assert_allclose(observed[0],[50.5,50.5,0.],atol=.5)
        np.testing.assert_array_equal(observed.sum(axis=1),[101,10000])
        np.testing.assert_array_equal(observed[1],[0,0,10000])
        np.testing.assert_allclose(artifacts["attribution"]["population"].sum(axis=0),sections[AGE_COLUMNS].sum())

    def test_cross_ccaa_flows_preserve_each_origin_not_destination_budget(self):
        from run_categorization import aggregate_scenario
        from categorization_core import AGE_COLUMNS
        sections=pd.DataFrame({"sb":["X","Y"],"pop_total":[10.,20.]})
        for c in AGE_COLUMNS:sections[c]=[1.,2.]
        stores=pd.DataFrame({"sb_modelo":["X","Y"],"ccaa_modelo":["CCAA1","CCAA2"]})
        matrix=sparse.csr_matrix([[.1,.9],[.8,.2]])
        result=aggregate_scenario(matrix,sections,stores,{"budgets":pd.Series([100,200],index=["X","Y"])},[1.,2.,3.])
        np.testing.assert_array_equal(result["flow_cents"].toarray(),[[10,90],[160,40]])
        np.testing.assert_array_equal(result["potential_cents"],[170,130])
        np.testing.assert_allclose(result["population"].sum(axis=0),[3.]*5)


if __name__ == "__main__":
    unittest.main()
