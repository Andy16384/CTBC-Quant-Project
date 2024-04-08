#This project is aim to

1. Connect Excel VBA to C++ and Quantlib, including debug both VBA and C++ code.
2. Interpolating the volatility surface(cubic on moneyness, linear on tenor).
3. Calibrating LMUV model onto the surface.

#Important files

1. FXvol.xlsm: Excel file, all VBA code are in this file.
2. Debug/Example.dll: the dynamic linking file.
3. Example.sln: Visual Studio Project Solution, Quantlib required.
4. Example/Example.cpp: C++ source code.
