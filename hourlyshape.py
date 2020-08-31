import requests

def calcShape(
    live_mode,
    country,
    deldate,
    delhours,
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

    #   country code: FR
    #   deldate:      date of the delivery of electricity
    #   delhours:     hour of the delivery of electricity
    #   quantitymw:   quantity of the delivery
    #   bucketrank:   bucket
    #   pxbucket:     price of the bucket
    #   bidaskmonth:   date of the bid-ask quotes
    #   bid_offer_PK/bid_offer_OP

    data = {
        'live_mode': live_mode,
        'country': country,
        'dates': deldate,
        'hours': delhours,
        'mw': quantitymw,
        'px_datetime': bucketrank,
        'px': pxbucket,
        'months': bidaskmonth,
        'bid_offer_PK': bid_offer_PK,
        'bid_offer_OP': bid_offer_OP,
        'gamma': gamma,
        'beta': beta,
        'holidayCostPerMWh': holidayCostPerMWh,
    }
    r = requests.post(
        'http://guru.trading.gdfsuez.net/hr_prof/price_hr_volume',
        json=data,
        verify=False,
    )
    return r.json()
