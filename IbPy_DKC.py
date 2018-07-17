import pandas as pd
import time
from IBWrapper import IBWrapper, contract
from ib.ext.EClientSocket import EClientSocket

desired_width = 320
pd.set_option('display.width', desired_width)


# Query functions -----------------------------------------------------------------
class IB:
    def portfolio(self):

        callback.updatePortfolio
        time.sleep(5)
        dataframe_portfolio = pd.DataFrame(callback.update_Portfolio,
                                           columns=['Contract ID', 'Currency', 'Expiry', 'Include Expired', 'Local Symbol', 'Multiplier', 'Primary Exchange', 'Right',
                                                    'Security Type', 'Strike', 'Symbol', 'Trading Class', 'Position', 'Market Price', 'Market Value', 'Average Cost',
                                                    'Unrealised PnL', 'Realised PnL', 'Account Name'])
        dataframe_portfolio[['Position', 'Market Price', 'Market Value', 'Average Cost', 'Unrealised PnL', 'Realised PnL']] = dataframe_portfolio[
            ['Position', 'Market Price', 'Market Value', 'Average Cost', 'Unrealised PnL', 'Realised PnL']].apply(pd.to_numeric)
        if not dataframe_portfolio.empty:
            print("\n######################### Portfolio ###########################")
            print(dataframe_portfolio)
            print("###############################################################\n")
        else:
            print("################### No portfolio to show ##################\n")

        return dataframe_portfolio

    def order(self, symbol, stype, exchange, currency, accountName, market, quantity, signal):
        tws.reqIds(1)
        order_id = callback.next_ValidId + 4
        time.sleep(10)
        iquantity = int(quantity)
        contract_info = create.create_contract(symbol, stype, exchange, currency)
        order_info = create.create_order(accountName, market, iquantity, signal)
        tws.placeOrder(order_id, contract_info, order_info)
        callback.next_ValidId += 5
        return

    def orderlist(self):
        callback.open_Order[:1]
        time.sleep(10)
        df_order_status = pd.DataFrame(callback.order_Status,
                                       columns=['orderId', 'status', 'filled', 'remaining', 'avgFillPrice',
                                                'permId', 'parentId', 'lastFillPrice', 'clientId', 'whyHeld'])
        if not df_order_status.empty:
            print("\n######################### Orders ##############################")
            print(df_order_status)
            print("###############################################################\n")
        else:
            print("\n###################### No orders to show #######################\n")

        return df_order_status

# Establish connection ----------------------------------------------------------------------
#accountName = "Your Account Name Here"
callback = IBWrapper()  # Instantiate IBWrapper. callback
tws = EClientSocket(callback)  # Instantiate EClientSocket and return data to callback
host = ""
port = 4002
clientId = 8
tws.eConnect(host, port, clientId)  # Connect to TWS
create = contract()  # Instantiate contract class
callback.initiate_variables()
tws.reqAccountUpdates(1, accountName)

# Portfolio Query ------------------------------------------------------------

ib = IB()
df_portfolio = ib.portfolio()
count = len(df_portfolio)

# Disconnect ------------------------------------------------------------------
time.sleep(2)
tws.isConnected()
tws.eDisconnect()
