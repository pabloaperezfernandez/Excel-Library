Attribute VB_Name = "PortfolioStatistics"
Option Explicit
Option Base 1

Function PortfolioSharpeRatio(ExcessReturnVector) As Double
'   Returns the Portfolio Sharpe Ratio
    PortfolioSharpeRatio = Application.Average(ExcessReturnVector) / Application.StDev(ExcessReturnVector)
End Function

Function PortfolioTreynorRatio(ExcessReturnVector, xsmktvec) As Double
'   Returns the Portfolio Treynor Ratio
'   Uses PortfolioBeta fn
    PortfolioTreynorRatio = Application.Average(ExcessReturnVector) / Application.Slope(ExcessReturnVector, xsmktvec)
End Function

Function PortfolioAlpha(ExcessReturnVector, xsmktvec) As Double
'   Returns the Portfolio Alpha (Intercept from OLS Regression)
    PortfolioAlpha = Application.Intercept(ExcessReturnVector, xsmktvec)
End Function

Function PortfolioAppraisalRatio(ExcessReturnVector, xsmktvec) As Double
'   Returns the Portfolio Appraisal Ratio
'   Uses PortfolioAlpha fn
'   Uses PortfolioSpecificRisk fn
    PortfolioAppraisalRatio = Application.Intercept(ExcessReturnVector, xsmktvec) / PortfolioSpecificRisk(ExcessReturnVector, xsmktvec, 1)
End Function

Function PortfolioSpecificRisk(ExcessReturnVector, xsmktvec, rperyr) As Double
'   Returns the Portfolio Specific Risk (OLS Standard Error of Y Estimate)
   PortfolioSpecificRisk = Application.StEyx(ExcessReturnVector, xsmktvec) * Sqr(rperyr)
End Function

Function M2RAP(rf, rp, sp, rm, sm) As Double
'   Returns the RAP measure (Modigliani & Modigliani, 1997)
    M2RAP = rf + (sm / sp) * (rp - rf)
End Function

Function M2RAPA(rf, rp, sp, rm, sm) As Double
'   Returns the RAP measure (Modigliani & Modigliani, 1997)
'   Uses M2RAP fn
    Let M2RAPA = M2RAP(rf, rp, sp, rm, sm) - rf
End Function
