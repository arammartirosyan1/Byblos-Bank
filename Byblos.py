import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta

ls_person = ["Գույքի սեփականատեր", "", "Հասցե մարզ", "Հասցե ամբողջական", " Անձնագրի համար", "Երբ է տրվել", "Ում կողմից", " Ծննդյան ամսաթիվ", "Սոցիալական քարտի համար", "Հեռախոս"]

ls_property = ["Ապահովագրվող գույքի անվանում", "", "Հասցե մարզ", "Հասցե ամբողջական", "Սկիզբ", "Ավարտ", "Գույքի արժեք", "Ապահովա-գրական գումար", "Սակագին (տոկոս)", "Ապահովա-գրավճար"]

ls_borrower = ["Վարկառու", "", "Հասցե մարզ", "Հասցե ամբողջական", " Անձնագրի համար", "Երբ է տրվել", "Ում կողմից", " Ծննդյան ամսաթիվ", "Սոցիալական քարտի համար", "Հեռախոս"]

ls_city = ['Երևան', 'Արմավիր', 'Արարատ', 'Արագածոտն', 'Կոտայք', 'Շիրակ', 'Լոռի', 'Տավուշ', 'Գեղարքունիք',
           'Վայոց Ձոր', 'Սյունիք']
ls_2 = ["ք․", "ք.", "Ք․", "Ք.", "քաղաք", "Քաղաք", "մ․", "մ.", "Մ․", "Մ.", "մարզ", "Մարզ", "ՀՀ", "հհ"]
ls_1 = {}


#Stexcel nor excel filer amen toxi hamar anrandzin
df = pd.read_excel("Byblos.xlsx", header=None)

df_1 = df.loc[4:5]
df_2 = df.loc[9:10]
df_3 = df.loc[15:16]

len_row = df_1.columns
for i in range(len(df_1)):
    data = df_1.values[i]
    new = {"": data}
    exc_data = pd.DataFrame(new, index=ls_person)
    exc_data.to_excel('News/Insured.xlsx')

for i in range(len(df_2)):
    data = df_2.values[i]
    new = {"": data}
    exc_data = pd.DataFrame(new, index=ls_borrower)
    exc_data.to_excel('News/Owner.xlsx')

for i in range(len(df_3)):
    data = df_3.values[i]
    new = {"": data}
    exc_data = pd.DataFrame(new, index=ls_property)
    exc_data.to_excel('News/Property.xlsx')


######################################################
#Stexcel apahovadirir dashtery
data_insured = df.values[5]
data_owner = df.values[10]
data_property = df.values[15]

if 'Բիբլոս' in str(df.values[2][2]):
    data_insured = pd.read_json('insured.json', orient='index')[0]
    x = str(df.values[10][8]).strip()
    ls_1["PROPERTY_OWNER_PERSON_SOCIAL_CARD"] = x
    ls_1["PROPERTY_OWNER_PERSON_PERS_BPR_USE"] = "1"
else:
    ls_1["IS_INSURED_PHYSICAL"] = "1"
    for i in range(len(data_insured)):
        data = data_insured
        try:
            if i == 0:
                x = str(data[i]).strip().split()
                ls_1["INSURED_NAME"] = x[1]
                ls_1["INSURED_LAST_NAME"] = x[0]
                ls_1["INSURED_SECOND_NAME"] = x[2]
        except:
            print("Error", data[i])


        try:
            if i == 2:
                x = str(data[i]).strip()
                if x == "Տավուշ":
                    ls_1["INSURED_REG_REGION"] = x
                    ls_1['INSURED_REG_CITY'] = "Այլ"
                else:
                    ls_1["INSURED_REG_REGION"] = x
                    ls_1['INSURED_REG_CITY'] = x
        except:
            print("error", data[i])

        try:
            if i == 3:
                new_data = str(data[i]).replace(',', "").strip()
                new_data = new_data.split(" ")
                for k in new_data[:3]:
                    for j in ls_city:
                        if j in k:
                            new_data_1 = str(data[i]).replace(k, '')
                            for q in ls_2:
                                if q in data[i]:
                                    qaxaq = q
                                    new_data_1 = new_data_1.replace(qaxaq, "").strip(',').strip().strip(',').strip()
                                    ls_1["INSURED_REG_COUNTRY"] = "ARM"
                                    ls_1['INSURED_REG_FULL_ADDRESS'] = new_data_1


        except:
            new_data_1 = str(new_data_1).strip().strip(",").strip()
            ls_1["INSURED_REG_COUNTRY"] = "ARM"
            ls_1['INSURED_REG_FULL_ADDRESS'] = new_data_1

        try:
            if i == 4:
                x = str(data[i]).strip()
                ls_1['INSURED_CITIZENSHIP'] = "ՀՀ"
                ls_1["INSURED_PASSPORT_NUMBER"] = x.replace('՛', '')
        except:
            print("error",  data[i])

        try:
            if i == 5:
                x = str(data[i]).strip()
                ls_1["INSURED_PASSPORT_ISSUE_DATE"] = x.replace('/', '.')
                start = datetime.datetime.strptime(x, "%d/%m/%Y")
                end = start + relativedelta(years=+10)
                ls_1["INSURED_PASSPORT_EXPIRY_DATE"] = end.strftime("%d.%m.%Y")

        except:
            print("error",  data[i])

        try:
            if i == 6:
                x = str(data[i]).strip()
                ls_1["INSURED_PASSPORT_AUTHORITY"] = x
        except:
            print("error",  data[i])

        try:
            if i == 7:
                x = str(data[i]).split()[0]
                x = datetime.datetime.strptime(x, "%d/%m/%Y").strftime("%d.%m.%Y")
                ls_1["INSURED_BIRTHDAY"] = x
        except:
            print("Error", data[i])

        try:
            if i == 8:
                x = str(data[i]).strip('.').strip('').split('.')
                ls_1["INSURED_SOCIAL_CARD"] = x[0]
                if int(x[0][:2]) > 50:
                    ls_1["INSURED_GENDER"] = "F"
                else:
                    ls_1["INSURED_GENDER"] = "M"
        except:
            print("Error", data[i])

        try:
            if i == 9:
                x = str(data[i]).split('.')[0].strip()
                if len(x) == 9:
                    ls_1["INSURED_MOBILE_PHONE"] = x
                else:
                    x = "0" + x[::-1][:8][::-1]
                    ls_1["INSURED_MOBILE_PHONE"] = x
        except:
            print("Error", data[i])

        ls_1["INSURED_MAIL"] = str(df.values[6][7]).strip()

    for i in range(len(data_owner)):
        data = data_owner
        if data_owner[0] == data_insured[0]:
            ls_1["IS_OWNER_PERSON_SAME_INSURED"] = "1"
        else:
            ls_1["PROPERTY_OWNER_PERSON_SOCIAL_CARD"] = str(data_insured[8]).strip()
            ls_1["PROPERTY_OWNER_PERSON_PERS_BPR_USE"] = "1"

        l = pd.Series(ls_1)
        l.to_json('News/Insured.json', indent=2, force_ascii=False)

####################################################################
#Stecum enq ararkayi jsony

for i in range(len(data_property)):
    data = data_property
    try:
        if i == 0:
            x = str(data[i]).strip()
            ls_1["PROPERTY_NAME"] = x
            ls_1["PROPERTY_TYPE"] = x
    except:
        print("Error", data[i])

    try:
        if i == 2:
            x = str(data[i]).strip()
            if x == "Տավուշ":
                ls_1["PROPERTY_REGION"] = x
                ls_1['PROPERTY_CITY'] = "Այլ"
            else:
                ls_1["PROPERTY_REGION"] = x
                ls_1['PROPERTY_CITY'] = x
    except:
        print("error", data[i])

    try:
        if i == 3:
            new_data = str(data[i]).replace(',', "").strip()
            new_data = new_data.split(" ")
            for k in new_data[:3]:
                for j in ls_city:
                    if j in k:
                        new_data_1 = str(data[i]).replace(k, '')
                        for q in ls_2:
                            if q in data[i]:
                                qaxaq = q
                                new_data_1 = new_data_1.replace(qaxaq, "").strip(',').strip().strip(',').strip()
                                ls_1["PROPERTY_COUNTRY"] = "ARM"
                                ls_1['PROPERTY_FULL_ADDRESS'] = new_data_1


    except:
        new_data_1 = str(new_data_1).strip().strip(",").strip()
        ls_1["PROPERTY_COUNTRY"] = "ARM"
        ls_1['PROPERTY_FULL_ADDRESS'] = new_data_1


    try:
        if i == 4:
            x = str(data[i]).strip().split(' ')[0]
            x = datetime.datetime.strptime(x, "%Y-%m-%d").strftime("%d.%m.%Y")
            ls_1["POLICY_FROM_DATE"] = x
            start_date = datetime.date.today().strftime("%d.%m.%Y")
            ls_1["POLICY_CREATION_DATE"] = start_date
    except:
        start_date = datetime.date.today()
        start = start_date.strftime("%d.%m.%Y")
        ls_1["POLICY_FROM_DATE"] = start
        ls_1["POLICY_CREATION_DATE"] = start


    try:
        if i == 5:
            x = str(data[i]).strip().split(' ')[0]
            x = datetime.datetime.strptime(x, "%Y-%m-%d").strftime("%d.%m.%Y")
            ls_1["POLICY_TO_DATE"] = x
    except:
        start_date = datetime.date.today()
        end_date = start_date + datetime.timedelta(days=365)
        end = end_date.strftime("%d.%m.%Y")
        ls_1["POLICY_TO_DATE"] = end


    try:
        if i == 6:
            x = str(data[i]).strip()
            ls_1["PROPERTY_MARKET_PRICE"] = x
    except:
        print("error", i, data[i])


    try:
        if i == 7:
            x = str(data[i]).strip()
            ls_1["PROPERTY_RISK_AMOUNT"] = x
            ls_1["PROPERTY_SUM_INSURED"] = x
            ls_1["POLICY_AMOUNT_CURRENCY"] = "AMD"
    except:
        print("error", i, data[i])

    try:
        if i == 8:
            x = str(data[i]).strip()
            ls_1["PROPERTY_INSURANCE_RATE"] = x
    except:
        print("error", i, data[i])


    try:
        if i == 9:
            x = str(data[i]).strip().split('.')[0]
            ls_1["PROPERTY_RISK_PREMIUM"] = x
            payment = str(x), start_date
            ls_1['POLICY_PAYMENT_SCHEDULE'] = ", ".join(payment)
    except:
        print("error", i, data[i])

    l = pd.Series(ls_1)
    l.to_json('News/Property.json', indent=2, force_ascii=False)

#############################################################

benef_data = pd.read_json('beneficiar.json', orient='index')[0]

excel_data = pd.read_json('News/Property.json', orient='index')[0]

agent_data = pd.read_json('agent.json', orient='index')[0]

if 'Բիբլոս' in str(df.values[2][2]):
    insurd_data = pd.read_json('insured.json', orient='index')[0]
    result = pd.concat([insurd_data, benef_data, excel_data, agent_data])
else:
    result = pd.concat([benef_data, excel_data, agent_data])

result.to_json('NEWS/New_Format.json', indent=2, force_ascii=False)




