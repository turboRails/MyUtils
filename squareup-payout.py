from square.client import Client
import sys
import pandas as pd
from datetime import datetime as dt, timezone,timedelta

#======================================================================#
#                振込日の期間  Period for Deposit date  (JST)           # 
start_date = "2022/06/30" 
end_date   = "2022/7/1"
#======================================================================#

sheet1 = "入金詳細"
columns = ['入金日', '決済日', '種類', '決済ID', '支払いID', '決済額', '手数料', '入金', '店舗']

def translate(name_en):
    name_jp = {'charge': '通常取引', 'refund': '払戻し'}.get(name_en)
    return name_jp if name_jp else name_en

# API Token for each Shop 
# Generated at developer.squareup.com after login as shop staff
token_list =[
            'EmZ9kSe6zY1p####### masked ####################EmZ9kSe6zY1pigme1',
            'EmZ9kSe6zY1p####### masked ####################EmZ9kSe6zY1pigme2',
            ]

start_time = dt.strptime(start_date, '%Y/%m/%d').strftime('%Y-%m-%d')
end_time   = (dt.strptime(end_date,   '%Y/%m/%d') + timedelta(hours=24)).strftime('%Y-%m-%d')
filename   = "CPJ" + start_date.replace('/','-')+ "_" + end_date.replace('/','-')  + ".xlsx"
err        = "got processing error. Please check on web"

rows     = []
writer   = pd.ExcelWriter(filename, engine='xlsxwriter')
workbook = writer.book
n_format = workbook.add_format({'num_format': '¥#,##0;[Red]-¥#,##0'})
d_format = workbook.add_format({'num_format': 'yyyy/mm/dd'})

class ResultError(Exception):
    def __init__(self, salary, message="Squareup API failed. Check Access Token"):
        for error in result.errors:
            print(error['category'])
            print(error['code'])
            print(error['detail'])
        self.result = result
        super().__init__(error['detail'])
        sys.exit(1)

for access_token in token_list:
    client = Client(
        access_token=access_token,
        environment='production',
        max_retries=2,
        timeout=125
        )

    # Get Shop name and id
    result = client.locations.list_locations()
    if result.is_success():
        shop        = result.body['locations'][0]['name']
        location_id = result.body['locations'][0]['id']
    else:
        ResultError(result)
        
    # Get all payouts in given period (Weekly bank transactions)
    results = client.payouts.list_payouts(
        location_id = location_id,
        begin_time  = start_time + "T00:00:00+09:00",
        end_time    = end_time   + "T00:00:00+09:00",
        sort_order  = "DESC"
    )
    payouts = results.body.get('payouts')
    if not payouts:
        print(f"{shop} {err}")
        continue
        
    # Get all payout entries in each payout
    for payout in payouts:
        try:
            entries = client.payouts.list_payout_entries(
                payout_id = payout['id']
            ).body.get('payout_entries')

            deposit_ = payout['created_at'].split("T")[0]
            deposit_date = dt.strptime(deposit_, '%Y-%m-%d').replace(tzinfo=timezone.utc).astimezone(tz=None).strftime('%Y/%m/%d')
        except KeyError:
            print(f"{shop} {err}")
            continue
            
        # Get payment details for all payments in one payout
        for entry in entries:
            try:
                pay_type = entry['type'].lower()
                payment_id = entry.get('type_' + pay_type + '_details').get('payment_id')
                if payment_id:
                    order_id = (client.payments.get_payment(payment_id = payment_id).body['payment']).get('order_id')
                
                row = []
                row.append(deposit_date)
                row.append(
                    dt.strptime(entry['effective_at'], '%Y-%m-%dT%H:%M:%SZ').replace(tzinfo=timezone.utc).astimezone(tz=None).strftime('%Y/%m/%d')
                )    
                row.append(translate(pay_type))
                row.append(order_id)
                row.append(payment_id)
                row.append(entry['gross_amount_money']['amount'])
                row.append((-1) * entry['fee_amount_money']['amount'])
                row.append(entry['net_amount_money']['amount'])
                row.append(shop)
            except KeyError:
                print(f"{shop} {err}")
                continue
            rows.append(row)
df = pd.DataFrame(rows, columns=columns)
df.to_excel(writer, sheet_name=sheet1, index=False)
worksheet = writer.sheets[sheet1]
worksheet.set_column('F:H', 10, n_format)
worksheet.set_column('A:B', 10, d_format)
writer.save()
