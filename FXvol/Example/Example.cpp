#pragma warning(disable:4244)
#pragma warning(disable:4819)

#include <ql/qldefines.hpp>
#if !defined(BOOST_ALL_NO_LIB) && defined(BOOST_MSVC)
#  include <ql/auto_link.hpp>
#endif

#include <iostream>
#include <vector>
#include <OAIdl.h>
#include <ql/time/calendars/japan.hpp>
#include <ql/termstructures/volatility/equityfx/blackconstantvol.hpp>
#include <ql/termstructures/yield/discountcurve.hpp>
#include <ql/termstructures/yieldtermstructure.hpp>
#include <ql/experimental/fx/blackdeltacalculator.hpp>
#include <ql/pricingengines/blackcalculator.hpp>
#include <ql/instruments/impliedvolatility.hpp>
#include <ql/math/interpolations/loginterpolation.hpp>
#include <ql/math/interpolations/linearinterpolation.hpp>
#include <ql/math/interpolations/cubicinterpolation.hpp>
#include <ql/math/solvers1d/brent.hpp>
#include <ql/math/optimization/levenbergmarquardt.hpp>
#include <ql/math/optimization/simplex.hpp>
#include <ql/math/optimization/constraint.hpp>
#include <ql/quotes/simplequote.hpp>
#include <boost/math/distributions.hpp>

#define cols(V) V.parray[0].rgsabound[0].cElements
#define rows(V) V.parray[0].rgsabound[1].cElements
#define M(V,i,j) ((((tagVARIANT*)(*(V.parray)).pvData))[rows(V) * j + i]).dblVal

using namespace QuantLib;

#pragma pack(4)
struct Interp
{
    VARIANT X_Y_;
    double x;
};
struct RateForDelta
{
    VARIANT Rate_;
    Volatility vol_;
    Real input_delta;
    Real Spot;
};
struct TotalData {
    //input
    VARIANT Vol_Table_Rd;
    VARIANT Vol_Table_Rf;
    VARIANT Vol_Table_;
    VARIANT Conventions_;
    int Today_;
    Real Spot_;
    int Expiry_;
    Real Strike_;

    //output

};
struct RI
{
    RateForDelta R;//Vol is included here.
    Interp I;

};
struct FXvol_Process_Type {
    
    //input
    VARIANT Vol_Table_Rd;
    VARIANT Vol_Table_Rf;
    VARIANT Vol_Table_;
    VARIANT Conventions_;
    Real Today_;
    Real Spot_;
    Integer ExpiryCase;
    
    //process
    VARIANT Rate_;
    VARIANT X_Y_;
    VARIANT Result;//deltaP, vol, strike.

    //Outcome;
    VARIANT Outcome;//vol, strike, premium.
} ;
struct LMUV_Calibration_Type {

    //Input
    Real Spot_;
    VARIANT Rate;
    VARIANT Premium;
    VARIANT Vol;
    VARIANT Strike;
    bool Ext_Alpha_;

    //Process
    VARIANT processVol;

    //Output
    Real p_;
    VARIANT Alpha;
    VARIANT Vol1;
    VARIANT Vol2;

};
#pragma pack()

Real __stdcall CubicSplineDLL(Interp* DeltaVol)
{
    int cols = cols(DeltaVol->X_Y_);// DeltaVol->X_Y_.parray[0].rgsabound[0].cElements;
    int rows = rows(DeltaVol->X_Y_);// DeltaVol->X_Y_.parray[0].rgsabound[1].cElements;
    double x = DeltaVol->x;
    std::vector < Real > deltaVec(cols), volVec(cols);
    for (int j = 0; j < cols; j++) {
        deltaVec[j] = M(DeltaVol->X_Y_, 0, j);//((((tagVARIANT*)(*(DeltaVol->X_Y_.parray)).pvData))[i * rows]).dblVal;
        volVec[j]   = M(DeltaVol->X_Y_, 1, j);//((((tagVARIANT*)(*(DeltaVol->X_Y_.parray)).pvData))[i * rows+1]).dblVal;
    }
    Real Lslope = (volVec[1] - volVec[0]) / (deltaVec[1] - deltaVec[0]);
    Real Rslope = (volVec[cols-1] - volVec[cols-2]) / (deltaVec[cols - 1] - deltaVec[cols - 2]);
    CubicInterpolation Firstbd(deltaVec.begin(), deltaVec.end(), volVec.begin(),
        CubicInterpolation::Spline, false,
        CubicInterpolation::FirstDerivative, Lslope,
        CubicInterpolation::FirstDerivative, Rslope);
    if (deltaVec[0] <= x && x <= deltaVec[deltaVec.size() - 1]) {
        return Firstbd(x);
    }
    else if (x < deltaVec[0]) {
        return volVec[0] + Lslope * (x - deltaVec[0]);
    }
    else {
        return volVec[deltaVec.size() - 1] + Rslope * (x - deltaVec[deltaVec.size() - 1]);
    }
}
Real __stdcall DeltaCP_DLL(RateForDelta* R){

    // input Call delta (positive) output Put dalta, input Put delta (negative) output Call dalta
    int cols = R->Rate_.parray[0].rgsabound[0].cElements;
    int rows = R->Rate_.parray[0].rgsabound[1].cElements;
    Real  rd = ((((tagVARIANT*)(*(R->Rate_.parray)).pvData))[2]).dblVal;
    Real  rf = ((((tagVARIANT*)(*(R->Rate_.parray)).pvData))[4]).dblVal;
    Real   t = ((((tagVARIANT*)(*(R->Rate_.parray)).pvData))[3]).dblVal;
    Option::Type ot_in = R->input_delta >  0 ? Option::Call : Option::Put;
    Option::Type otout = R->input_delta <= 0 ? Option::Call : Option::Put;

    DeltaVolQuote::DeltaType dt = t > 1 ? DeltaVolQuote::DeltaType::PaFwd : DeltaVolQuote::DeltaType::PaSpot;
    BlackDeltaCalculator input(ot_in, dt, R->Spot, std::exp(-rd * t), std::exp(-rf * t), R->vol_ * std::sqrt(t));
    BlackDeltaCalculator output(input);
    output.setOptionType(otout);
    output.setDeltaType(DeltaVolQuote::DeltaType::PaSpot);

    return output.deltaFromStrike(input.strikeFromDelta(R->input_delta));
}
Real __stdcall AtmDeltaDLL(RateForDelta* R) {

    //Although input RateForDelta, the input_delta is not used here

    int cols = R->Rate_.parray[0].rgsabound[0].cElements;
    int rows = R->Rate_.parray[0].rgsabound[1].cElements;
    Real  rd = M(R->Rate_, 0, 1);//((((tagVARIANT*)(*(R->Rate_.parray)).pvData))[2]).dblVal;
    Real  rf = M(R->Rate_, 0, 2);//((((tagVARIANT*)(*(R->Rate_.parray)).pvData))[4]).dblVal;
    Real   t = M(R->Rate_, 1, 1);//.parray)).pvData))[3]).dblVal;

    Option::Type ot = Option::Put;
    DeltaVolQuote::DeltaType dt = t > 1 ? DeltaVolQuote::DeltaType::PaFwd : 
                                          DeltaVolQuote::DeltaType::PaSpot;
    DeltaVolQuote::AtmType atmT = t > 1 ? DeltaVolQuote::AtmType::AtmFwd : 
                                          DeltaVolQuote::AtmType::AtmDeltaNeutral;
    BlackDeltaCalculator input(ot, dt, R->Spot, std::exp(-rd * t), 
                                                std::exp(-rf * t), R->vol_ * std::sqrt(t));
    Real atmStrike = input.atmStrike(atmT);
    input.setDeltaType(DeltaVolQuote::DeltaType::PaSpot);
    Real ttt = 1;// (t > 1) ? exp(-rf * t) : 1;
    return input.deltaFromStrike(atmStrike);
    
}
Real __stdcall FindVolDLL(TotalData* R) {
    
    std::vector < Date > dates_d, dates_f;
    std::vector < DiscountFactor > dfsd, dfsf;
    Calendar cal = Japan();
    DayCounter dc = Actual365Fixed();
    int rows_d = R->Vol_Table_Rd.parray[0].rgsabound[1].cElements;
    int rows_f = R->Vol_Table_Rf.parray[0].rgsabound[1].cElements;

    Date temp(((((tagVARIANT*)(*(R->Vol_Table_Rd.parray)).pvData))[rows_d]).dblVal);
    DiscountFactor D = 1.0;
    dates_d.push_back(temp);    dfsd.push_back(D);
    dates_f.push_back(temp);    dfsf.push_back(D);
    for (int i = 0; i < rows_d; i++) {
        Date temp(((((tagVARIANT*)(*(R->Vol_Table_Rd.parray)).pvData))[3 * rows_d + i]).dblVal);
        D = ((((tagVARIANT*)(*(R->Vol_Table_Rd.parray)).pvData))[5 * rows_d + i]).dblVal;
        dates_d.push_back(temp);
        dfsd.push_back(D);
    }

    for (int i = 0; i < rows_f; i++) {
        Date temp(Date::serial_type(((((tagVARIANT*)(*(R->Vol_Table_Rf.parray)).pvData))[3 * rows_f + i]).dblVal));
        DiscountFactor D = ((((tagVARIANT*)(*(R->Vol_Table_Rf.parray)).pvData))[5 * rows_f + i]).dblVal;
        dates_f.push_back(temp);
        dfsf.push_back(D);
    }
    InterpolatedDiscountCurve < LogLinear > dDiscount(dates_d, dfsd, dc, cal);
    InterpolatedDiscountCurve < LogLinear > fDiscount(dates_f, dfsf, dc, cal);
    //Date ttemp(Date::serial_type(((((tagVARIANT*)(*(R->Vol_Table_Rf.parray)).pvData))[3 * 27 + 5]).dblVal));
    //ttemp = ttemp + 5 * Months;

    return 1.0;//fDiscount.discount(ttemp);
}

void __stdcall FxVolProcessDLL(FXvol_Process_Type* FX) {

    //Discounts    
    Calendar calendar = Japan();
    DayCounter dc = Actual365Fixed();
    std::vector < Date > dates_d, dates_f;
    std::vector < DiscountFactor > dfsd, dfsf;
    Date temp(FX->Today_);
    dates_d.push_back(temp);    dfsd.push_back(1.0);
    dates_f.push_back(temp);    dfsf.push_back(1.0);

    for (unsigned int i = 0; i < rows(FX->Vol_Table_Rd); i++) {
        Date temp(M(FX->Vol_Table_Rd, i, 3));
        dates_d.push_back(temp);
        dfsd.push_back(M(FX->Vol_Table_Rd, i, 5));
    }
    for (unsigned int i = 0; i < rows(FX->Vol_Table_Rf); i++) {
        Date temp(M(FX->Vol_Table_Rf, i, 3));
        dates_f.push_back(temp);
        dfsf.push_back(M(FX->Vol_Table_Rf, i, 5));
    }
    InterpolatedDiscountCurve < LogLinear > dDiscount(dates_d, dfsd, dc, calendar);
    InterpolatedDiscountCurve < LogLinear > fDiscount(dates_f, dfsf, dc, calendar);

    Date TodaysDate(FX->Today_);
    Date Expiry(M(FX->Vol_Table_, FX->ExpiryCase, 0));
    bool longterm = M(FX->Rate_, 1, 1) > 1 ? true : false;
    Integer fixingDays = 2;
    Date settlementDate = calendar.advance(TodaysDate, fixingDays, Days);
    Date es = calendar.advance(Expiry, fixingDays, Days);
    // must be business days
    settlementDate = calendar.adjust(settlementDate);
    es = calendar.adjust(es);
    M(FX->Rate_, 0, 0) = FX->Today_;
    M(FX->Rate_, 1, 0) = M(FX->Vol_Table_, FX->ExpiryCase, 0);
    M(FX->Rate_, 1, 1) = (M(FX->Rate_, 1, 0) - M(FX->Rate_, 0, 0)) / 365;
    M(FX->Rate_, 0, 1) = -log(dDiscount.discount(es) / dDiscount.discount(settlementDate)) / M(FX->Rate_, 1, 1);//Need some adjustment
    M(FX->Rate_, 0, 2) = -log(fDiscount.discount(es) / fDiscount.discount(settlementDate)) / M(FX->Rate_, 1, 1);//Need some adjustment
    M(FX->Rate_, 1, 2) = FX->Spot_ * exp((M(FX->Rate_, 0, 1) - M(FX->Rate_, 0, 2)) * M(FX->Rate_, 1, 1));

    M(FX->X_Y_, 1, 0) = M(FX->Vol_Table_, FX->ExpiryCase, 4) + M(FX->Vol_Table_, FX->ExpiryCase, 10) + M(FX->Vol_Table_, FX->ExpiryCase, 8) / 2;
    M(FX->X_Y_, 1, 1) = M(FX->Vol_Table_, FX->ExpiryCase, 4) + M(FX->Vol_Table_, FX->ExpiryCase, 9) + M(FX->Vol_Table_, FX->ExpiryCase, 7) / 2;
    M(FX->X_Y_, 1, 2) = M(FX->Vol_Table_, FX->ExpiryCase, 4);
    M(FX->X_Y_, 1, 3) = M(FX->Vol_Table_, FX->ExpiryCase, 4) + M(FX->Vol_Table_, FX->ExpiryCase, 9) - M(FX->Vol_Table_, FX->ExpiryCase, 7) / 2;
    M(FX->X_Y_, 1, 4) = M(FX->Vol_Table_, FX->ExpiryCase, 4) + M(FX->Vol_Table_, FX->ExpiryCase, 10) - M(FX->Vol_Table_, FX->ExpiryCase, 8) / 2;

    RateForDelta ratefordelta;
    ratefordelta.Rate_ = FX->Rate_;
    ratefordelta.Spot = FX->Spot_;
    ratefordelta.vol_ = M(FX->X_Y_, 1, 0) / 100;
    ratefordelta.input_delta = 0.1;
    M(FX->X_Y_, 0, 0) = DeltaCP_DLL(&ratefordelta);
    ratefordelta.vol_ = M(FX->X_Y_, 1, 1) / 100;
    ratefordelta.input_delta = 0.25;
    M(FX->X_Y_, 0, 1) = DeltaCP_DLL(&ratefordelta);
    ratefordelta.vol_ = M(FX->X_Y_, 1, 2) / 100;
    M(FX->X_Y_, 0, 2) = AtmDeltaDLL(&ratefordelta);
    M(FX->X_Y_, 0, 3) = M(FX->Rate_, 1, 1) > 1 ? -0.25 * exp(-M(FX->Rate_, 0, 2) * M(FX->Rate_, 1, 1)) : -0.25;
    M(FX->X_Y_, 0, 4) = M(FX->Rate_, 1, 1) > 1 ? -0.1 * exp(-M(FX->Rate_, 0, 2) * M(FX->Rate_, 1, 1)) : -0.1;

    Interp IT;
    IT.X_Y_ = FX->X_Y_;

    //Call 0~8    //0->0.05    //1->0.1
    for (Real DeltaPillar = 0; DeltaPillar < 9; DeltaPillar++) {
        Real diff = 1.0;
        ratefordelta.input_delta = 0.05 * (DeltaPillar + 1);
        IT.x = ratefordelta.input_delta - 1;

        while (abs(diff) > 1E-12) {
            diff = DeltaCP_DLL(&ratefordelta) - IT.x;
            IT.x = DeltaCP_DLL(&ratefordelta);
            ratefordelta.vol_ = CubicSplineDLL(&IT) / 100;
        }
        //Put Delta
        M(FX->Result, 0, int(DeltaPillar)) = IT.x;
        //Vol
        M(FX->Result, 1, int(DeltaPillar)) = ratefordelta.vol_;
        M(FX->Outcome, 0, int(DeltaPillar)) = M(FX->Result, 1, int(DeltaPillar)) * 100;
        //Strike
        M(FX->Result, 2, int(DeltaPillar)) = M(FX->Rate_, 1, 1) > 1.0 ?
            (ratefordelta.input_delta * exp(-M(FX->Rate_, 0, 2) * M(FX->Rate_, 1, 1)) - IT.x) * FX->Spot_ * exp(M(FX->Rate_, 0, 1) * M(FX->Rate_, 1, 1)) :
            (ratefordelta.input_delta - IT.x) * FX->Spot_ * exp(M(FX->Rate_, 0, 1) * M(FX->Rate_, 1, 1));
        M(FX->Outcome, 1, int(DeltaPillar)) = M(FX->Result, 2, int(DeltaPillar));

        //Premium
        BlackCalculator black(Option::Call,
            M(FX->Outcome, 1, int(DeltaPillar)),
            M(FX->Rate_, 1, 2), M(FX->Result, 1, int(DeltaPillar)) * sqrt(M(FX->Rate_, 1, 1)),
            exp(-M(FX->Rate_, 0, 1) * M(FX->Rate_, 1, 1)));
        M(FX->Outcome, 2, int(DeltaPillar)) = black.value();//Need adjustment

    }

    //Put Delta
    M(FX->Result, 0, 9) = M(FX->X_Y_, 0, 2);
    //Vol
    M(FX->Result, 1, 9) = M(FX->X_Y_, 1, 2) / 100;
    M(FX->Outcome, 0, 9) = M(FX->X_Y_, 1, 2);
    //Strike
    M(FX->Result, 2, 9) = M(FX->Rate_, 1, 1) > 1 ? M(FX->Rate_, 1, 2) :
        M(FX->Rate_, 1, 2) * exp(-M(FX->Result, 1, 9) * M(FX->Result, 1, 9) / 2 * M(FX->Rate_, 1, 1));
    M(FX->Outcome, 1, 9) = M(FX->Result, 2, 9);
    //Premium
    BlackCalculator black(Option::Call,
        M(FX->Outcome, 1, 9),
        M(FX->Rate_, 1, 2), M(FX->Result, 1, 9) * sqrt(M(FX->Rate_, 1, 1)),
        exp(-M(FX->Rate_, 0, 1) * M(FX->Rate_, 1, 1)));
    M(FX->Outcome, 2, 9) = black.value();//Need adjustment


    //Put 10~19    //10 ->-0.45    //11->-0.4
    for (Real DeltaPillar = 10; DeltaPillar < cols(FX->Result); DeltaPillar++) {
        //Put Delta
        M(FX->Result, 0, int(DeltaPillar)) = 0.05 * (DeltaPillar - 19) * ((M(FX->Rate_, 1, 1) > 1) ? exp(-M(FX->Rate_, 0, 2) * M(FX->Rate_, 1, 1)) : 1);

        //Vol
        IT.x = M(FX->Result, 0, int(DeltaPillar));
        M(FX->Result, 1, int(DeltaPillar)) = CubicSplineDLL(&IT) / 100;
        M(FX->Outcome, 0, int(DeltaPillar)) = M(FX->Result, 1, int(DeltaPillar)) * 100;

        //Strike
        ratefordelta.input_delta = 0.05 * (DeltaPillar - 19);
        ratefordelta.vol_ = M(FX->Result, 1, int(DeltaPillar));
        M(FX->Result, 2, int(DeltaPillar)) = (DeltaCP_DLL(&ratefordelta) - IT.x) * FX->Spot_ * exp(M(FX->Rate_, 0, 1) * M(FX->Rate_, 1, 1));
        M(FX->Outcome, 1, int(DeltaPillar)) = M(FX->Result, 2, int(DeltaPillar));

        //Premium
        BlackCalculator black(Option::Put,
            M(FX->Outcome, 1, int(DeltaPillar)),
            M(FX->Rate_, 1, 2),
            M(FX->Result, 1, int(DeltaPillar)) * sqrt(M(FX->Rate_, 1, 1)),
            exp(-M(FX->Rate_, 0, 1) * M(FX->Rate_, 1, 1)));
        M(FX->Outcome, 2, int(DeltaPillar)) = black.value();//Need adjustment
    }
}

class LMUVCaliProblemFunction1 : public CostFunction {
public:
    LMUVCaliProblemFunction1(LMUV_Calibration_Type* lc) : LC(lc) {}
    /*
    Real value(const Array& x) const { // x = [p1, Vol1, Vol2, Alpha1, Alpha2 ] 
        Array tmpRes = values(x);
        Real res = 0;
        for (size_t i = 0; i < 55; i++) res += tmpRes[i] * tmpRes[i];
        return sqrt(res);
    }*/
    Array values(const Array& x) const { // x = {p1, Vol1, Vol2, Alpha1, Alpha2 } 
        Array res(55);
        for (int i = 0; i < 55; i++) {
            Option::Type type = i < 33 ? Option::Call : Option::Put;
            BlackCalculator black1(type, M(LC->Strike, i, 0) - x[3] * exp((M(LC->Rate, 11, 0) - M(LC->Rate, 11, 1)) * M(LC->Rate, 11, 2)), (LC->Spot_ - x[3]) * exp((M(LC->Rate, 11, 0) - M(LC->Rate, 11, 1)) * M(LC->Rate, 11, 2)), x[1] * sqrt(M(LC->Rate, 11, 2)), exp(-M(LC->Rate, 11, 0) * M(LC->Rate, 11, 2)));
            BlackCalculator black2(type, M(LC->Strike, i, 0) - x[4] * exp((M(LC->Rate, 11, 0) - M(LC->Rate, 11, 1)) * M(LC->Rate, 11, 2)), (LC->Spot_ - x[4]) * exp((M(LC->Rate, 11, 0) - M(LC->Rate, 11, 1)) * M(LC->Rate, 11, 2)), x[2] * sqrt(M(LC->Rate, 11, 2)), exp(-M(LC->Rate, 11, 0) * M(LC->Rate, 11, 2)));
            res[i] = (1 - x[0]) * black1.value() + x[0] * black2.value() - M(LC->Premium, i, 0);
        }
        return res;
    }
private:
    LMUV_Calibration_Type* LC;
};

void __stdcall LMUVCalibrationDLL(LMUV_Calibration_Type* LC) {
    LMUVCaliProblemFunction1 optFunction(LC);
    EndCriteria myEndCrit(1000, 100, 1e-5, 1e-5, 1e-5);
    Array startVal = { 0.05,0.1,0.5,0,0 };
    NonhomogeneousBoundaryConstraint constraint({ 0,0,0,-100,-100 }, { 0.5,5,5,100,100 });
    Problem myProb(optFunction, constraint, startVal);
    Simplex solver(0.1);
    EndCriteria::Type solvedCrit = solver.minimize(myProb, myEndCrit);

    LC->p_ = myProb.currentValue()[0];
    M(LC->Alpha, 0, 0) = myProb.currentValue()[3];
    M(LC->Alpha, 0, 1) = myProb.currentValue()[4];
}