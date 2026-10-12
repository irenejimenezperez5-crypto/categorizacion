"""DEGURBA scope and distance-neutral rural allocation over the frozen population.

Keep original age attribution as an inspectable demographic input. Rural weights
are averaged within municipality and SB of location, preserving every SB total.
"""
import numpy as np
import pandas as pd

def rural_allocation_weights(points):
    result=points.copy()
    deg=pd.to_numeric(result['degurba'],errors='raise').fillna(0).astype(int)
    if not deg.isin([0,1,2,3]).all():
        raise ValueError('DEGURBA no válido; no asumir ámbito urbano para un valor desconocido')
    demand=pd.to_numeric(result['weighted_population'],errors='raise')
    if not np.isfinite(demand).all() or (demand<0).any():
        raise ValueError('Demanda no finita o negativa')
    result['geographic_scope']=np.select([deg.isin([1,2]),deg.eq(3)],['No rural','Rural'],default='Sin resolver')
    result['visit_eligible']=deg.isin([1,2])
    result['allocation_weight']=demand
    result['allocation_method']=np.select([deg.isin([1,2]),deg.eq(3)],
       ['No rural: demanda espacial y cercanía','Rural: reparto igual municipio–SB'],
       default='Ámbito sin resolver: demanda provisional; revisar antes de visitar')
    result['rural_group_key']=''
    result['rural_group_count']=0
    result['rural_group_demand']=0.
    rural=deg.eq(3)
    municipalities=result.loc[rural,'municipality_code'].fillna('').astype(str)
    known=municipalities.str.fullmatch(r'\d{5}')
    keys=result.loc[rural,'sb_code'].astype(str)+'|'+municipalities.where(known,'SIN_MUNICIPIO')
    result.loc[rural,'rural_group_key']=keys
    result.loc[rural,'allocation_weight']=result.loc[rural].groupby(keys)['weighted_population'].transform('mean')
    result.loc[rural,'rural_group_count']=result.loc[rural].groupby(keys)['weighted_population'].transform('size').astype(int)
    result.loc[rural,'rural_group_demand']=result.loc[rural].groupby(keys)['weighted_population'].transform('sum')
    result.loc[keys[~known].index,'allocation_method']='Rural provisional: reparto igual de puntos rurales sin municipio dentro del SB'
    result['sb_allocation_weight']=result.groupby('sb_code')['allocation_weight'].transform('sum')
    if (result.sb_allocation_weight<=0).any():
        raise ValueError('SB sin demanda; requiere imputación identificada antes de repartir')
    result['allocation_share']=result.allocation_weight/result.sb_allocation_weight
    before=result.groupby('sb_code').weighted_population.sum()
    after=result.groupby('sb_code').allocation_weight.sum()
    if not np.allclose(before,after,rtol=1e-12,atol=1e-8):
        raise AssertionError('El reparto rural cambió demanda total por SB')
    return result
