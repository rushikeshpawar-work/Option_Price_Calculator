import time
import schedule
import pandas as pd 
from pandas import ExcelFile
from datetime import datetime
from datetime import date
import math
from scipy.stats import norm
import sklearn
from sklearn.metrics import mean_squared_error
from sklearn.metrics import r2_score
from sklearn.metrics import mean_absolute_error
import matplotlib.pyplot as plt
import xlwings as xw

#==================================================================================
# Input are taken from user just once time

lower_limit = int(input("Enter minimum price of strick price (Ex.: 17500) : "))
upper_limit = int(input("Enter maximun price of strick price (Ex.: 18300) : "))

now = datetime.now()
current_time = now.strftime("%H:%M")

# following job funcion will execute after every one minut 
def job():

    # ============================= Imorting the Data =============================

    option_workbook = xw.Book('D:\my_project\live_market_data.xlsx')
    data = pd.read_excel(r'D:\my_project\live_market_data.xlsx')
    option_data = data

    index_workbook = xw.Book('D:\my_project\index_data.xlsx')
    index_data = pd.read_excel(r'D:\my_project\index_data.xlsx')
    nifty50 = index_data.iloc[0,3]

    option_workbook.save()
    index_workbook.save()

    option_data = option_data.iloc[: , [0,1,5,6,8,10,9,11,13,24,25,27,29,28,30,32]]

    # ============================= Data Transformation =============================

    dateformat = "%d-%b-%Y"
    current_date = date.today()
    current_date = current_date.strftime(dateformat)

    option_data = option_data[(option_data["expiryDate"] >= current_date) & (option_data["strikePrice"] >= lower_limit) & (option_data["strikePrice"] <= upper_limit )] 

    option_data['year_to_maturity'] = (((pd.to_datetime(option_data.loc[:,'expiryDate']) - (pd.to_datetime(current_date))).dt.days)/365)

    call_df = option_data.iloc[: , [2,0,1,3,6,16,5,7,8]]
    put_df = option_data.iloc[: , [9,0,1,10,13,16,12,14,15]]

    # function to calculate the price of option
    def black_scholes(S, K, T, sigma,  r=0.1,option_type='call'):
        """
        Calculate the price of a call or put option using the Black-Scholes model.
        
        Parameters:
        S (float): underlying asset price(current index value)
        K (float): option strike price
        T (float): time to maturity (in years)
        r (float): risk-free interest rate
        sigma (float): volatility of the underlying asset (in %)
        option_type (str): type of option, either 'call' or 'put'
        
        Returns:
        float: option price
        """

        if (sigma != 0 and T != 0): 
            d1 = (math.log(S/K) + (r + 0.5*sigma**2)*T) / (sigma*math.sqrt(T))
            d2 = d1 - sigma*math.sqrt(T)
            
            if option_type == 'call':
                price = S*norm.cdf(d1) - K*math.exp(-r*T)*norm.cdf(d2)
            elif option_type == 'put':
                price = K*math.exp(-r*T)*norm.cdf(-d2) - S*norm.cdf(-d1)
            else:
                raise ValueError("Option type must be either 'call' or 'put'")
            return price
        else:
            return None


    call_df['p_implicit_volatility'] = ((call_df.loc[:,'CE.impliedVolatility']) / 100)
    put_df['p_implicit_volatility'] = ((put_df.loc[:,'PE.impliedVolatility']) / 100)

    call_df['CE.estimeted_price'] = call_df.apply(lambda x: black_scholes(nifty50, x['strikePrice'], x['year_to_maturity'], x['p_implicit_volatility'] ,0.1,'call'), axis=1)

    call_df['CE.p_difference'] = call_df.apply(lambda x: (((x['CE.lastPrice'] - x['CE.estimeted_price']) / x['CE.estimeted_price']) * 100) if x['CE.lastPrice'] > 0 and x['CE.estimeted_price'] != None else None, axis=1)

    put_df['PE.estimeted_price'] = put_df.apply(lambda x: black_scholes(nifty50, x['strikePrice'], x['year_to_maturity'], x['p_implicit_volatility'], 0.1,'put'), axis=1)

    put_df['PE.p_difference'] = put_df.apply(lambda x: (((x['PE.lastPrice'] - x['PE.estimeted_price']) / x['PE.estimeted_price']) * 100) if x['PE.lastPrice'] > 0 and x['PE.estimeted_price'] != None else None, axis=1)

    call_put_df = pd.concat([call_df, put_df], axis=1)

    df5 = option_data.iloc[:,[3,5,6,8,10,12,13,15]]

    call_put_df = pd.concat([call_put_df, df5], axis=1)

    output_df = call_put_df.iloc[:,[2,0,3,6,4,7,8,10,11,1,12,15,18,16,19,20,22,23]]

    # ============================= Model Evaluation =============================

    PE_evaluation = put_df.iloc[:,[7,10]]
    PE_evaluation = PE_evaluation[(PE_evaluation['PE.lastPrice'] != 0)]
    PE_evaluation = PE_evaluation.dropna()

    PE_dic={"Parameters":"Put"}
    PE_dic.update({'MSE' : (mean_squared_error(PE_evaluation['PE.lastPrice'],PE_evaluation['PE.estimeted_price']))})
    PE_dic.update({'MAE' : (mean_absolute_error(PE_evaluation['PE.lastPrice'],PE_evaluation['PE.estimeted_price']))})
    PE_dic.update({'R-square' : (r2_score(PE_evaluation['PE.lastPrice'],PE_evaluation['PE.estimeted_price']))})

    # creating data frame outof dictionary
    PE_evaluation_df = pd.DataFrame.from_dict(PE_dic, orient ='index') 

    CE_evaluation = call_df.iloc[:,[7,10]]
    CE_evaluation = CE_evaluation[(CE_evaluation['CE.lastPrice'] != 0)]
    CE_evaluation = CE_evaluation.dropna()

    CE_dic={"Parameters":"Call"}
    CE_dic.update({'MSE' : (mean_squared_error(CE_evaluation['CE.lastPrice'],CE_evaluation['CE.estimeted_price']))})
    CE_dic.update({'MAE' : (mean_absolute_error(CE_evaluation['CE.lastPrice'],CE_evaluation['CE.estimeted_price']))})
    CE_dic.update({'R-square' : (r2_score(CE_evaluation['CE.lastPrice'],CE_evaluation['CE.estimeted_price']))})

    # creating data frame outof dictionary 
    CE_evaluation_df = pd.DataFrame.from_dict(CE_dic, orient ='index')

    evaluation_df = pd.concat([CE_evaluation_df,PE_evaluation_df],axis=1)

    # ============================= Pltoing the chart =============================
    put_fig = plt.figure()
    pe_x = PE_evaluation["PE.lastPrice"]
    pe_y = PE_evaluation["PE.estimeted_price"]
    plt.plot(pe_x,linestyle = "-",color='black')
    plt.plot(pe_y,linestyle= "dotted",color='lawngreen')
    plt.grid(False)
    plt.legend(["Market price", 'Calculated price'],loc ="center",bbox_to_anchor =(0., 1.09, 1., .102), ncol = 2)
    plt.title("Put Actualprice vs Estimeted price")

    call_fig = plt.figure()
    ce_x = CE_evaluation["CE.lastPrice"]
    ce_y = CE_evaluation["CE.estimeted_price"]
    plt.plot(ce_x,linestyle = "-",color='black')
    plt.plot(ce_y,linestyle= "dotted",color='lawngreen')
    plt.legend(["Market price", 'Calculated price'],loc ="center",bbox_to_anchor =(0., 1.09, 1., .102), ncol = 2)
    plt.grid(False)
    plt.title("Call Actualprice vs Estimeted price")

    # ============================= loading the data into ".xlsx" format =============================

    file_path = 'D:\my_project\Black_scholes_output.xlsx'
    workbook = xw.Book('D:\my_project\Black_scholes_output.xlsx')

    # Check if the sheet is present
    sheet_name1 = 'Output'
    sheet_name2 = 'BSM_evaluation'
    sheet_name3 = 'Linechart'

    if sheet_name1 in [sheet.name for sheet in workbook.sheets]:
        # Delete the sheet
        workbook.sheets[sheet_name1].delete()
    if sheet_name2 in [sheet.name for sheet in workbook.sheets]:
        # Delete the sheet
        workbook.sheets[sheet_name2].delete()
    if sheet_name3 in [sheet.name for sheet in workbook.sheets]:
        # Delete the sheet
        workbook.sheets[sheet_name3].delete()
        # Create a new sheet
    workbook.sheets.add(sheet_name3)
    workbook.sheets.add(sheet_name2)
    workbook.sheets.add(sheet_name1)

    # Save the workbook
    workbook.save(file_path)

    # Open the workbook
    workbook = xw.Book(file_path)
    #selecting worksheet
    output_worksheet = workbook.sheets[sheet_name1]
    evaluation_worksheet = workbook.sheets[sheet_name2]
    linechart_worksheet = workbook.sheets[sheet_name3]
    # Insert the DataFrame in the worksheet
    output_worksheet.range('A1').options(index=False,header=2).value = output_df
    evaluation_worksheet.range('A1').options(index=True).value = evaluation_df

    workbook.save()

    #--------------------------inserting line chart-------------------------------------------
    
    sht = workbook.sheets['Linechart']
    sht.pictures.add(
        put_fig,
        name="Put_linechart",
        update=True,
        left=sht.range("P2").left,
        top=sht.range("P2").top,
        height=400,
        width=600,
    )

    workbook.save()

    sht.pictures.add(
        call_fig,
        name="Call_linechart",
        update=True,
        left=sht.range("B2").left,
        top=sht.range("B2").top,
        height=400,
        width=600,
    )

    workbook.save()

#============================== job scheduling ============================================


schedule.every(1).minutes.until("15:32").do(job)

while (current_time <= '15:32'):
     now = datetime.now()
     current_time = now.strftime("%H:%M")
     
     # to check any schedule task is pending or not
     schedule.run_pending()
     time.sleep(1)
