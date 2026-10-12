"""Compare economic segmentation with an independent exhaustive small-data oracle."""
import itertools
from pathlib import Path
import random
import sys
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from economic_breaks import natural_breaks, assign

def loss(groups):
    return sum(sum((x-sum(g)/len(g))**2 for x in g) for g in groups if g)

class EconomicPolicyTests(unittest.TestCase):
    def test_optimum_matches_exhaustive_partitions_with_unequal_duplicate_weights(self):
        rng=random.Random(2027)
        for _ in range(24):
            unique=sorted(rng.sample(range(100,10000),9))
            values=[v/100 for v in unique for _ in range(rng.randrange(1,6))]
            candidates=[]
            for split in itertools.combinations(range(1,len(unique)),3):
                breaks=[unique[i]/100 for i in split]
                groups=[[] for _ in range(4)]
                for v in values:groups[sum(v>=b for b in breaks)].append(v)
                candidates.append(loss(groups))
            result=natural_breaks(values)
            labels=assign(values,result['cuts_ascending_eur'])
            actual=loss([[v for v,c in zip(values,labels) if c==letter] for letter in 'DCBA'])
            self.assertAlmostEqual(actual,min(candidates),places=7)
            self.assertEqual(sum(result['class_counts_D_C_B_A']),len(values))

    def test_equal_values_are_never_split_and_lower_bounds_are_inclusive(self):
        values=[100]*7+[1000]*3+[5000]*2+[20000]*5
        result=natural_breaks(values)
        self.assertEqual(result['class_counts_D_C_B_A'],[7,3,2,5])
        c,b,a=result['cuts_ascending_eur']
        self.assertEqual(assign([c-.01,c,b-.01,b,a-.01,a],[c,b,a]).tolist(),list('DCCBBA'))

    def test_unsorted_input_does_not_change_cuts(self):
        values=[1,1,2,3,10,15,40,50,99]
        self.assertEqual(natural_breaks(values),natural_breaks(list(reversed(values))))

    def test_invalid_reference_is_rejected(self):
        for values in [[0,1,2,3,4],[1,2,float('nan'),4,5],[1,2,3,3]]:
            with self.assertRaises(ValueError):natural_breaks(values)

if __name__=='__main__':unittest.main()
