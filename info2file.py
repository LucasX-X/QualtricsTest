import cx_Oracle
import sys
import csv
from datetime import datetime
from datetime import timedelta

date = datetime.now() - timedelta(days=1)
date = date.strftime('%d-%b-%y')
date = date.upper()
print(date)

dsn_tns = cx_Oracle.makedsn('wyoming.rice.edu', '1521', service_name='dwdprod.rice.edu')
con = cx_Oracle.connect(user='WHFEPMGR', password='FEP4N4LYTICS', dsn=dsn_tns)
# con = cx_Oracle.connect(user='FEP_Python', password='Pyth0nF0rF3P', dsn=dsn_tns)

cur = con.cursor()

queryString1 = "SELECT REQOR_FIRST_NAME, REQOR_LAST_NAME, REQOR_EMAIL, REQ_STATEMENT_OF_WORK, REQ_ID "
queryString2 = "FROM V_FEP_WORK_ORDER_DETAILS WHERE REQ__WORK_COMPLETE_DT LIKE '{date}%' AND REQ_STATUS LIKE 'Complete' ".format(date=date)
queryString3 = "AND REQ_TYPE in ('Events', 'Elevators &' || ' Lifts', 'Carpentry', 'Electrical', 'Exterior Building', " \
               "'HVACR', 'Interior Building', 'Leaks', 'Mechanical', 'Painting', 'Plumbing', 'Vehicle &' || ' Equipment Repair', " \
               "'Grounds/Landscaping', 'Moving', 'Odor', 'Pest Control', 'Pest/Animal Control', 'Pest/AnimalControl', 'Custodial') "
queryString4 = "AND CREWNAME IN ('01 - Air Conditioning', '02 - Plumbing-Ext-Mech Repair', '03 - Electrical'," \
               "'04 - Carpentry - Painting', '07 - Preventive Maintenance', '08 - Elevators','11 - Grounds', '12 - Movers-Solid Waste', " \
               "'13 - Custodial South', '13 - Custodial East', '13 - Custodial North', '13 - Custodial Admin', " \
               "'13 - Custodial West','17 - Equipment Repair', '24 - Arboriculture')"
queryString5 = "AND CREWNAME IN ('08 - Elevators','11 - Grounds', '12 - Movers-Solid Waste', " \
               "'13 - Custodial South', '13 - Custodial East', '13 - Custodial North', '13 - Custodial Admin', " \
               "'13 - Custodial West','17 - Equipment Repair', '24 - Arboriculture')"

if int(date[0:2]) % 2 == 1:
    queryString = queryString1 + queryString2 + queryString3 + queryString4
else:
    queryString = queryString1 + queryString2 + queryString3 + queryString5

cur.execute(queryString)

# fetch the data and description
data = cur.fetchall()
fields = cur.description

cur.close()
con.close()

contact_list = []
# transfer it from table to json
for row in data:
    record = {}
    for i in range(len(fields)):
        record[fields[i][0]] = row[i]
    contact_list.append(record)

contact_list = contact_list

# change the content make it legal to be sent
t = []
for i in range(len(contact_list)):
    d = {}
    d['firstName'] = contact_list[i]['REQOR_FIRST_NAME']
    d['lastName'] = contact_list[i]['REQOR_LAST_NAME']
    d['email'] = contact_list[i]['REQOR_EMAIL']
    d['Statement of work'] = contact_list[i]['REQ_STATEMENT_OF_WORK']
    d['Work Order'] = contact_list[i]['REQ_ID']
    d['Date'] = date
    t.append(d)

contact_list = t[:]


with open('C:/Users/tlm6/Desktop/test.csv', 'a+') as csvfile:
    fieldnames = ['firstName', 'lastName', 'email', 'Statement of work', 'Work Order', 'Date']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(contact_list)
    csvfile.close()
