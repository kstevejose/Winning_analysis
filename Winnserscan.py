import pandas as pd
import numpy as np
import time

startTime = time.time()

executionTime = (time.time() - startTime)
print('Time to import modules ' + str(executionTime))

startTime_2 = time.time()
#####your python script#####

executionTime_2 = (time.time() - startTime_2)
print('Time to run the main Python script: ' + str(executionTime_2))

def winning_analysis():
    first_excel='MIDWEEK_bet_details_15 Sep 2021.xlsx'


    df=pd.read_excel(first_excel)

    df.columns = map(str.lower, df.columns)
    draft=df.apply(lambda x: x.astype(str).str.lower())
    draft['winning_amt']=draft['winning_amt'].astype(float)


    group_df=draft.groupby(['account_id','tsn','bet_datetime','panel_id','panel_amt','bettype_name', 'winning_amt']).agg({'bet_selection': lambda x : ','.join(x)})

    group_df.reset_index(inplace=True)
    # group_df['panel_amt']=group_df['panel_amt'].astype(int)
#
    final=group_df.groupby(['account_id','bettype_name','bet_selection','panel_amt'])[['winning_amt']].agg('sum')
    final2=group_df.groupby(['bettype_name','bet_selection', 'panel_amt'])[['winning_amt']].agg('sum')
    final.reset_index(inplace=True)
    final2.reset_index(inplace=True)
#
#
    df1=final[final['bettype_name']=='2 sure']
    df_1=final2[final2['bettype_name']=='2 sure ']
    df2=final[final['bettype_name']=='direct 3']
    df_2=final2[final2['bettype_name']=='direct 3']
    df3=final2[final2['bettype_name']=='perm 2']
    df4=final2[final2['bettype_name']=='perm 3']
    df5=final2[final2['bettype_name']=='first number drop']
#
    df1.sort_values(by='winning_amt', ascending=False, inplace=True)
    df_1.sort_values(by='winning_amt', ascending=False,inplace=True)
    df_2.sort_values(by='winning_amt', ascending=False,inplace=True)
    df2.sort_values(by='winning_amt', ascending=False, inplace=True )
    df3.sort_values(by='winning_amt', ascending=False, inplace=True)
    df4.sort_values(by='winning_amt', ascending=False, inplace=True)
    df5.sort_values(by='winning_amt', ascending=False, inplace=True)


    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('Bettypes.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    df1.to_excel(writer, sheet_name='2 sure', index=False)
    df_1.to_excel(writer, sheet_name='2 sure with account_id', index=False)
    df_2.to_excel(writer, sheet_name='direct 3', index=False)
    df2.to_excel(writer, sheet_name='direct 3 with account_id', index=False)
    df3.to_excel(writer, sheet_name='perm 2', index=False)
    df4.to_excel(writer, sheet_name='perm 3', index=False)
    df5.to_excel(writer, sheet_name='first number drop', index=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    print(df1.head(5))
    print(df_1.head(5))
    # print(df3.head(5))
    # print(df4.head(5))
    # print(df5.head(5))
    # print(final2.head(5))
#
winning_analysis()
# #final.to_excel('final.xlsx')


#print(df1.head(5))
