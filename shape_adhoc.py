import numpy as np
import pandas as pd
from Old import Dates
from hourlyshape import calcShape


def convertDatelistToString(deldate, usetime):

    molist = [0] * len(deldate)
    if usetime:
        myformat = '%Y-%m-%d %H:%M:%S'
    else:
        myformat = '%Y-%m-%d'

    for k, mo in enumerate(deldate):
        if isinstance(mo, (int, float, complex)):
            mo = pd.to_datetime(str(mo));
            molist[k] = Dates.fromexceldate(mo).strftime(myformat)
        else:
            mo = pd.to_datetime(str(mo));
            molist[k] = mo.strftime(myformat)

    return molist

def calc_shape_adhoc(
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

    deldate = convertDatelistToString(deldate, 0)
    bidaskmonth = convertDatelistToString(bidaskmonth, 0)
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
        ','.join(deldate),
        ','.join(delhour),
        ','.join(quantitymw),
        ','.join(bucketrank),
        ','.join(pxbucket),
        ','.join(bidaskmonth),
        ','.join(bid_offer_PK),
        ','.join(bid_offer_OP),
        gamma,
        beta,
        holidayCostPerMWh,
    )
    #if 'message' in resubase.keys():
    #    return resubase['message']
    return resubase

def weekend(df):
    df['WeekDay'] = df['Date'].dt.dayofweek;
    weekend = np.empty(len(df['Date']), dtype=object);
    for k in range(len(df['Date'])):
        if df['WeekDay'][k] >= 5 :
            weekend[k] = 'WeekEnd'
        else:
            weekend[k] = 'WorkingDay'
        df['WeekEnd'] = weekend;

def peak_offpeak(df) :
    peakoffpeak = np.empty(len(df['Date']), dtype=object);
    for k in range(len(df['Date'])):
        if df['WeekEnd'][k]=='WeekEnd' or df['Heure'][k]<=7 or df['Heure'][k]>=20:
            peakoffpeak[k] = 'OffPeak'
        else:
            peakoffpeak[k] = 'Peak'
        df['Peak/OffPeak'] = peakoffpeak;
    return 0

def dentelle(df):

    offpeak_mean = df.groupby('Peak/OffPeak')['CdC'].mean()['OffPeak'];
    peak_mean = df.groupby('Peak/OffPeak')['CdC'].mean()['Peak'];
    df.loc[df['Peak/OffPeak']=='OffPeak','Dentelle']=offpeak_mean;
    df.loc[df['Peak/OffPeak'] == 'Peak', 'Dentelle'] = peak_mean;
    df['Dentelle']=df['CdC']-df['Dentelle']
    return 0

def month(df):
    month = np.empty(len(df['Date']), dtype=object);
    for k in range(len(df['Date'])):
        month[k] = df['Date'][k].month
    df['Month'] = month;

def saison(df):
    saison = np.empty(len(df['Date']), dtype=object);
    for k in range(len(df['Date'])):
        if df['Month'][k] <= 3 or df['Month'][k] >= 10:
            saison[k] = 'Hiver'
        else:
            saison[k] = 'Ete'
        df['Saison'] = saison;

def result(
        file_name,
        sheet_name,
        gamma
):
    df = pd.read_excel(file_name, sheet_name=sheet_name);
    weekend(df)
    peak_offpeak(df);
    month(df);
    saison(df)
    dentelle(df);
    df_date = df['Date'];
    df_heure = df['Heure'];
    df_dentelle = df['Dentelle'];

    deldate = df_date.to_numpy();
    delhour = df_heure.to_numpy();
    quantitymw = df_dentelle.to_numpy();

    resubase = calc_shape_adhoc(
        live_mode=live_mode,
        country=country,
        deldate=deldate,
        delhour=delhour,
        quantitymw=quantitymw,
        bucketrank=bucketrank,
        pxbucket=pxbucket,
        bidaskmonth=bidaskmonth,
        bid_offer_PK=bid_offer_PK,
        bid_offer_OP=bid_offer_OP,
        gamma=gamma,
        beta=beta,
        holidayCostPerMWh=holidayCost
    );
    conso_total = df['CdC'].sum()
    brique = resubase['Bid/Ask (€)'] / conso_total;
    dentelle_pourcentage = (df[df['Dentelle'] > 0]['Dentelle'].sum() - df[df['Dentelle'] < 0]['Dentelle'].sum())/conso_total
    hiver_pourcentage = df[df['Saison']=='Hiver']['CdC'].sum()/conso_total;
    weekend_pourcentage = df[df['WeekEnd']=='WeekEnd']['CdC'].sum()/conso_total;
    offpeak_pourcentage = df[df['Peak/OffPeak']=='OffPeak']['CdC'].sum()/conso_total;

    return [brique, dentelle_pourcentage,hiver_pourcentage,weekend_pourcentage,offpeak_pourcentage]


live_mode = 1;
country = "FR";

beta = 0.34;
gammaN1 = 0.5;
gammaN2 = 0.55;
gammaN3 = 0.6;
gammaN3 = 0.65;
holidayCost = 0.2;

file_name = 'Test.xlsx'
sheet_name_1 = 'N+1'
sheet_name_2 = 'N+2'
sheet_name_3 = 'N+3'
sheet_name_4 = 'N+4'

df_bidask = pd.read_excel(file_name,sheet_name="BidAsk");
df_bidask_month = df_bidask['bidaskmonth'];
df_bidask_bidofferpk = df_bidask['bidofferpk'];

df_bidask_bidofferop = df_bidask['bidofferop'];

df_bucketrank = pd.read_excel(file_name,sheet_name="BucketRank");
df_bucketrank_date = df_bucketrank['bucketrank'];
df_bucketrank_price = df_bucketrank['pricebucket'];
df_bucketrank_pricedelta = df_bucketrank['pricebucket + dP'];

bucketrank = df_bucketrank_date.to_numpy();
pxbucket = df_bucketrank_price.to_numpy();
bidaskmonth = df_bidask_month.to_numpy();
bid_offer_PK = df_bidask_bidofferpk.to_numpy();
bid_offer_OP = df_bidask_bidofferop.to_numpy();

result_N1 = result(file_name=file_name, sheet_name=sheet_name_1,gamma=gammaN1)
result_N2 = result(file_name=file_name, sheet_name=sheet_name_2,gamma=gammaN2)
result_N3 = result(file_name=file_name, sheet_name=sheet_name_3,gamma=gammaN3)
result_N4 = result(file_name=file_name, sheet_name=sheet_name_4,gamma=gammaN3)


print("\nRésultats N+1 :")
print("Brique : ",result_N1[0])
print("% Dentelle :", result_N1[1])
print(result_N1)


print("\nRésultats N+2 :")
print("Brique : ",result_N2[0])
print("% Dentelle :", result_N2[1])
print(result_N2)

print("\nRésultats N+3 :")
print("Brique : ",result_N3[0])
print("% Dentelle :", result_N3[1])
print(result_N3)

print("\nRésultats N+4 :")
print("Brique : ",result_N4[0])
print("% Dentelle :", result_N4[1])
print(result_N4)