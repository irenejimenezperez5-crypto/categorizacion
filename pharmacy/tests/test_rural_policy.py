"""Business invariants for rural equality, budgets and known visit eligibility."""
from pathlib import Path
import sys
import unittest
import numpy as np
import pandas as pd
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from rural_allocation import rural_allocation_weights
from review_categorization import cents_allocation

class RuralPolicyTests(unittest.TestCase):
    def rows(self):
        return pd.DataFrame({'id':['r1','r2','city','other','unknown'],
         'sb_code':['010010','010010','010010','010020','010020'],
         'municipality_code':['01001','01001','01002','01001',''],
         'degurba':[3,3,1,3,0],'weighted_population':[100.,300.,600.,100.,100.]})

    def test_rural_pool_is_equal_without_flattening_urban_demand_or_other_sb(self):
        source=self.rows();r=rural_allocation_weights(source)
        self.assertEqual(r.allocation_weight.tolist(),[200.,200.,600.,100.,100.])
        self.assertEqual(r.rural_group_count.tolist(),[2,2,0,1,0])
        self.assertEqual(r.visit_eligible.tolist(),[False,False,True,False,False])
        self.assertTrue(source.equals(self.rows()))
        self.assertTrue(r.weighted_population.equals(source.weighted_population))

    def test_rural_euros_equal_to_one_cent_and_every_location_sb_conserved(self):
        r=rural_allocation_weights(self.rows())
        budgets=pd.Series({'010010':100.01,'010020':10.01})
        cents=cents_allocation(r.allocation_weight,budgets,r.sb_code)
        self.assertLessEqual(abs(cents.iloc[0]-cents.iloc[1]),1)
        self.assertEqual(cents.groupby(r.sb_code).sum().to_dict(),{'010010':10001,'010020':1001})
        self.assertAlmostEqual(r.loc[2,'allocation_share'],.6)

    def test_missing_municipality_uses_explicit_rural_sb_pool(self):
        source=self.rows();source.loc[:1,'municipality_code']=''
        r=rural_allocation_weights(source)
        self.assertTrue(r.loc[:1,'allocation_method'].str.contains('provisional').all())
        self.assertEqual(r.loc[:1,'allocation_weight'].tolist(),[200.,200.])
        self.assertEqual(r.loc[4,'geographic_scope'],'Sin resolver')

    def test_invalid_geography_and_demographic_inputs_are_rejected(self):
        for field,value in [('degurba',5),('weighted_population',-1.),('weighted_population',np.nan)]:
            source=self.rows();source.loc[0,field]=value
            with self.assertRaises(ValueError):rural_allocation_weights(source)
        source=self.rows();source.loc[source.sb_code.eq('010010'),'weighted_population']=0
        with self.assertRaises(ValueError):rural_allocation_weights(source)

if __name__=='__main__':unittest.main()
