
import numpy as np
import pandas as pd
import xlwings as xw
from Old.dmp import dmpdic, getdic
from Old import Dates
from hourlyshape import calcShape
import getinstance


@xw.sub
def get_workbook_name():
    wb = xw.Book.caller()
    wb.sheets['Sheet1'].range('D3').value = wb.name


@xw.arg('filename')
@xw.arg('mykey1')
@xw.arg('mykey2')
@xw.arg('mykey3')
@xw.ret(transpose=True)
def xl_getdic(filename, mykey1, mykey2='', mykey3=''):
    return getdic(filename, mykey1, mykey2, mykey3)


@xw.func
@xw.arg('pos', np.array, ndim=1)
@xw.arg('months', np.array, ndim=1)
@xw.arg('excel_evaldate')
@xw.arg('stratype')
@xw.arg('basepeak')
@xw.arg('midask', np.array, ndim=2)
@xw.arg('filename')
def xl_calc_cost_single(
    pos, months, excel_evaldate, stratype, basepeak, midask, filename,
):

    obj = getinstance(stratype)
    monthmidask = []
    quartermidask = []
    yearmidask = []
    for dta in midask:
        if dta[0][0] == 'M':
            monthmidask.append(float(dta[1]))
        elif dta[0][0] == 'Q':
            quartermidask.append(float(dta[1]))
        elif dta[0][0] == 'Y':
            yearmidask.append(float(dta[1]))
    obj.bidask[basepeak]['M'] = monthmidask
    obj.bidask[basepeak]['Q'] = quartermidask
    obj.bidask[basepeak]['Y'] = yearmidask
    if isinstance(excel_evaldate, (int, float, complex)):
        evaldate = Dates.fromexceldate(excel_evaldate)
    else:
        evaldate = excel_evaldate
    molist = [0] * len(months)

    for k, mo in enumerate(months):
        if isinstance(mo, (int, float, complex)):
            molist[k] = Dates.fromexceldate(mo)
        else:
            molist[k] = mo

    daypos = {'pos': pd.Series(pos), 'months': np.array(molist), 'evaldate': evaldate}
    resubase = obj.getHedge(evaldate, daypos, basepeak)
    dmpdic(resubase, filename)
    return 1


def convertDatelistToString(deldate, usetime):

    molist = [0] * len(deldate)
    if usetime:
        myformat = '%Y-%m-%d %H:%M:%S'
    else:
        myformat = '%Y-%m-%d'

    for k, mo in enumerate(deldate):
        if isinstance(mo, (int, float, complex)):
            molist[k] = Dates.fromexceldate(mo).strftime(myformat)
        else:
            molist[k] = mo.strftime(myformat)

    return molist


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
def xl_calc_shapehourlylocal(
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

    from Old.price_hr_vol import PriceHourlyVolume
    zz = PriceHourlyVolume()
    rr = {
        "live_mode": live_mode,        "country": country,
        "dates":','.join(molist), "hours": ','.join(delhour), "mw": ','.join(quantitymw),
        "px_datetime": ','.join(bucketrank),
        "px": ','.join(pxbucket),
        "months": ','.join(bidasklist),
        "bid_offer_PK":','.join(bid_offer_PK),
        "bid_offer_OP": ','.join(bid_offer_OP),
        "gamma": gamma,
        "beta": beta,
        "holidayCostPerMWh": holidayCostPerMWh
    }
    resubase = zz.calcShapelocal(rr)
    if 'message' in resubase.keys():
        return resubase['message']
    return resubase


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
def xl_calc_shapehourly(
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
