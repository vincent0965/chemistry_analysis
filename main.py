import os
import time
import random
import numpy as np
import pandas as pd
from tqdm import tqdm
from scipy.optimize import minimize, basinhopping
import win32com.client as win32


class Aspen(object):
    def __init__(self,
                 data_path,
                 SimulationName,
                 write_value_path,
                 output_path,
                 target_component,
                 other_component,
                 parameter_path,
                 sample_size,
                 parameter):

        self.data_path = data_path
        self.SimulationName = SimulationName
        self.write_value_path = write_value_path
        self.output_path = output_path
        self.target_component = target_component
        self.other_component = other_component
        self.parameter_path = parameter_path
        self.sample_size = sample_size
        self.parameter = parameter


    def Aspen_minimize(self):

        # read data
        xls = pd.ExcelFile(self.data_path)
        sheets = xls.sheet_names
        input_data = xls.parse(sheets[0])
        real_data = xls.parse(sheets[1])

        # pre process data
        input_data, input_component = self.preprocess_data(data = input_data)
        real_data, real_component = self.preprocess_data(data = real_data)

        # Open Aspen
        Application = self.Open_Aspen(SimulationName = self.SimulationName)
        sample = random.sample(range(input_data.shape[1]), self.sample_size)
        print(sample)
        # 求 parameter 最佳解

        minimizer_kwargs = dict(
            args = (Application, input_data, input_component, real_data,
                    real_component, self.target_component, self.other_component,
                    self.parameter_path, sample),
            tol = 1e-4,
            method = 'L-BFGS-B'
        )

        res = basinhopping(
            func = self.run_Aspen,
            x0 = self.parameter,
            minimizer_kwargs = minimizer_kwargs,
            niter = 10000,
            stepsize = 0.02,
            interval = 3,
            disp = True
        )

        # res = minimize(self.run_Aspen, self.parameter, args=(Application, input_data, input_component, real_data,
        #                                                      real_component, self.target_component, self.other_component,
        #                                                      self.parameter_path, sample),
        #                method = 'Nelder-Mead', options = {"maxiter": 10000}, tol = 1e-4)
        # basinhopping(self.run_Aspen, self.parameter, niter=100, minimizer_kwargs=minimizer_kwargs)
        # method = "L-BFGS-B", options = {"maxiter": 10, "disp": True, 'eps': 0.0001}, tol = 1e-4)

        # 求 parameter最佳解之數據模擬
        MAPE, result = self.run_Aspen(res.x, Application, input_data, input_component, real_data, real_component,
                                      self.target_component, self.other_component, self.parameter_path, sample, df = True)

        return MAPE, res.x, result


    def preprocess_data(self, data):
        rows_name = data["Unnamed: 0"].to_list()
        data.index = rows_name
        data.drop(columns = "Unnamed: 0", inplace = True)

        return data, rows_name


    def Open_Aspen(self, SimulationName):
        Application = win32.Dispatch("Apwn.Document.37.0")
        # 打開三效蒸發器文件
        Application.InitFromArchive2(os.path.abspath(SimulationName))
        # 設置AP用戶界面的可見性，1為可見，0為不可見
        Application.Visible = 0

        return Application


    def write_value_into_Aspen(self, Application, write_value_path, data, input_component, run_numders):

        columns_names = data.columns.to_list()
        for i in range(len(input_component)):
            if i == 0:
                Application.Tree.FindNode("/Data/Blocks/3F8011/Input/" + input_component[i]).Value = data.loc[input_component[i], columns_names[run_numders]]
            elif i == 1:
                Application.Tree.FindNode("/Data/Blocks/3F8011/Input/" + input_component[i]).Value = data.loc[input_component[i], columns_names[run_numders]]
            else:
                Application.Tree.FindNode(write_value_path + input_component[i]).Value = data.loc[input_component[i], columns_names[run_numders]]

        return Application


    def show_Aspen_finish_data(self, Application, data, output_path, real_component, run_numders):

        pred_component_value_list = []
        for i in real_component:
            pred_component_value_list.append(Application.Tree.FindNode(output_path + i).Value)

        finish_data = pd.DataFrame({"real":data[data.columns[run_numders]], "pred":pred_component_value_list})

        return Application, finish_data


    def mape(self, y_true, y_pred):

        return np.mean(np.abs((y_true - y_pred) / y_true)) * 100


    def run_Aspen(self, parameter, Application, input_data, input_component, real_data, real_component,
                  target_component, other_component, parameter_path, sample, df = False):

        # Fill parameter to Aspen
        print([round(num, 4) for num in parameter])
        for i in range(len(parameter_path)):
            Application.Tree.FindNode("/Data/Reactions/Reactions/R-1/Input/PRE_EXP/" + parameter_path[i]).Value = parameter[2*i]
            Application.Tree.FindNode("/Data/Reactions/Reactions/R-1/Input/ACT_ENERGY/" + parameter_path[i]).Value = parameter[2*i+1]

        MAPE = 0
        counts =0
        for i in tqdm(sample):
            # Fill data to Aspen
            Application = self.write_value_into_Aspen(Application = Application,
                                                      write_value_path = self.write_value_path,
                                                      data = input_data,
                                                      input_component = input_component,
                                                      run_numders = i)
            # Run Aspen
            Application.Engine.Run2(1)
            time.sleep(7)

            Application, finish_data = self.show_Aspen_finish_data(Application = Application,
                                                                   data = real_data,
                                                                   output_path = self.output_path,
                                                                   real_component = real_component,
                                                                   run_numders = i)

            if counts == 0:
                finish_real_data = finish_data["real"]
                finish_pred_data = finish_data["pred"]
            else:
                finish_real_data  = pd.concat([finish_real_data, finish_data["real"]], axis = 1)
                finish_pred_data = pd.concat([finish_pred_data, finish_data["pred"]], axis = 1)

            for j in target_component:
                single_MAPE = self.mape(y_true = finish_data.loc[j, "real"], y_pred = finish_data.loc[j, "pred"])
                print("data case:", i, "component:", j, "MAPE:", round(single_MAPE, 4))
                MAPE += single_MAPE

            for j in other_component:
                other_MAPE = self.mape(y_true = finish_data.loc[j, "real"], y_pred = finish_data.loc[j, "pred"])
                print("data case:", i, "component:", j, "MAPE:", round(other_MAPE, 4))

            counts += 1

        for j in target_component + other_component:
            Total_MAPE = self.mape(y_true = np.array(finish_real_data.loc[j,:]), y_pred = np.array(finish_pred_data.loc[j,:]))
            print("component:", j, "Total_MAPE", round(Total_MAPE, 4))

        MAPE = MAPE / len(sample) / len(target_component)
        print(round(MAPE, 4))
        if df == False:
            return MAPE
        else:
            return MAPE, finish_data


if __name__ == "__main__":
    data_path = "R_Auto/R_Auto_3.xlsx"                            # data 路徑(xlsx檔)
    SimulationName = "R_Auto/Recycle Try_Auto.bkp"              # Aspen 路徑(bkp檔)
    write_value_path = r"/Data/Streams/S8-103/Input/FLOW/MIXED/"     # S8-103 path
    output_path = r"/Data/Streams/S8-105/Output/MASSFLOW/MIXED/"     # S8-105 path
    parameter_path = ["5", "6", "7", "8", "11", "12", "15", "16", "18", "28", "30", "34", "35"]           # parameter 路徑
    parameter = [2.7151, 10.2306, 2.0807, 10.2331, 0.5714, 9.3881, 3.7073, 9.3729, 5.6353, 10.2429, 15.9908,
                 10.2331, 182.35, 38.5618, 0.0325, 16.7327, 0.0073, 27.1173, 0.1028, 28.3811, 0.1573, 27.4092, 28.7562, 70.1263, 0.0, 50.7412]      # parameter
    target_component = ["BZ", "EZB", "TOL"]                         # target component mape
    other_component = ["OX", "MX", "PX"]                            # other component mape
    sample_size = 64                                                # 抽樣的樣本數
    X = Aspen(data_path = data_path,
              SimulationName = SimulationName,
              write_value_path = write_value_path,
              output_path = output_path,
              target_component = target_component,
              other_component = other_component,
              parameter_path = parameter_path,
              sample_size = sample_size,
              parameter = parameter)
    print(X.Aspen_minimize())

