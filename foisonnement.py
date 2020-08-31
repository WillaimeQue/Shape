from hourlyshape import calcShape
import numpy as np
import xlwings as xw
import pandas as pd

@xw.func
@xw.arg('live_mode')
@xw.arg('country')
@xw.arg('deldate', np.array, ndim=1)
@xw.arg('delhour', np.array, ndim=1)
@xw.arg('quantitymw', np.array, ndim=1)
@xw.arg('bucketrank', np.array, ndim=1)
@xw.arg('pxbucket', np.array, ndim=1)
@xw.arg('bidaskmonth', np.array, ndim=1)
@xw.arg('bid_offer_PK', np.array, ndim=1)
@xw.arg('bid_offer_OP', np.array, ndim=1)
@xw.arg('gamma')
@xw.arg('beta')
@xw.arg('holidayCostPerMWh')
def xl_calc_shapehourly_fs(
    live_mode,
    country,
    deldate,
    delhour,
    quantitymw,
    bucketrank,
    pxbucket,
    bidaskmonth,
    bid_offer_PK,
    bid_offer_OP,
    gamma,
    beta,
    holidayCostPerMWh,
):

    molist = convertDatelistToString(deldate, 0)
    bidasklist = convertDatelistToString(bidaskmonth, 0)
    bucketrank = convertDatelistToString(bucketrank, 1)
    delhour = [str(delhourcpt) for delhourcpt in delhour]
    quantitymw = [str(quantitymwcpt) for quantitymwcpt in quantitymw]
    bucketrank = [str(bucketrankcpt) for bucketrankcpt in bucketrank]
    pxbucket = [str(pxbucketcpt) for pxbucketcpt in pxbucket]
    bid_offer_PK = [str(bid_offer_PKcpt) for bid_offer_PKcpt in bid_offer_PK]
    bid_offer_OP = [str(bid_offer_OPcpt) for bid_offer_OPcpt in bid_offer_OP]

    resubase = calcShape(
        live_mode,
        country,
        ','.join(molist),
        ','.join(delhour),
        ','.join(quantitymw),
        ','.join(bucketrank),
        ','.join(pxbucket),
        ','.join(bidasklist),
        ','.join(bid_offer_PK),
        ','.join(bid_offer_OP),
        gamma,
        beta,
        holidayCostPerMWh,
    )
    if 'message' in resubase.keys():
        return resubase['message']
    return resubase


if __name__ == '__main__':
    # To run this with the debug server, set UDF_DEBUG_SERVER = True in the xlwings VBA module
    xw.serve()
    # from udf import xl_calc_cost_single
    # import numpy as np
    # import pandas as pd