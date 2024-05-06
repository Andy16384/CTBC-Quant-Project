# Project's aim

- Connect Excel VBA to C++ and Quantlib, including debug both VBA and C++ code.
- Interpolating the volatility surface(cubic on moneyness, linear on tenor).
- Calibrating LMUV model onto the surface.

# Important files

- FXvol.xlsm: Excel file, all VBA code are in this file.
- Debug/Example.dll: the dynamic linking file.
- Example.sln: Visual Studio Project Solution, Quantlib required.
- Example/Example.cpp: C++ source code.

# Some Pictures
![alt text](https://github.com/Andy16384/CTBC-Quant-Project/blob/main/Fixed%20Income%20Derivatives/Tree_Setting_and_Result.png)
![alt text](https://github.com/Andy16384/CTBC-Quant-Project/blob/main/Fixed%20Income%20Derivatives/Vol_and_Rate_Data.png)
