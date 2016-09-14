query_for_gvia_month_report = '''SELECT
  NameTeam AS [צוות],
  NameCustomer AS [שם לקוח],
  MiddlePay AS [אמצעי תשלום],
  AgreementKind AS [סוג לקוח],
  CASE WHEN proceFinished=0 THEN 'כן' ELSE '' END AS [לקוח משפטי],
  CASE WHEN Conditions='דולר' THEN ISNULL(PayDollar, 0)*4 ELSE ISNULL(PayDollar, 0) END AS [תשלום ייעוץ חודשי],
  ISNULL(SumBonus,0) as [תשלום על בונוס],
  ISNULL(SumSpecial,0) as [תשלומים מיוחדים],
  ISNULL(NumOfDebt,0) as [מספר התשלומים מעבר לאשראי],
  ISNULL(Debt,0) as [סכום התשלומים מעבר לאשראי],
  ISNULL(Remark,'') as [הערות],
  ISNULL(Text, '') AS [הערות מגיליון הגביה]
FROM
  dbo.tblCustomers
  LEFT JOIN
  dbo.tblTeamList
    ON dbo.tblCustomers.KodTeamCare = dbo.tblTeamList.KodTeamCare
  LEFT JOIN
  (SELECT * FROM
    (
      SELECT
        IDagreement,
        KodAgreementKind,
        KodCustomer,
        PayProcess,
        PayDollar,
        agreementFinish,
        row_number() OVER(PARTITION BY KodCustomer ORDER BY DateNew DESC) AS orderAgreementByDateDesc
      FROM
        dbo.tblAgreementConditionAdvice
    ) AS temp
      WHERE temp.orderAgreementByDateDesc = 1) AS tblLastAgreements
    ON dbo.tblCustomers.KodCustomer = tblLastAgreements.KodCustomer
  LEFT JOIN tblJudicialProc
    ON tblCustomers.KodCustomer = tblJudicialProc.kodCustomer
  LEFT JOIN dbo.tblMiddlePay
    ON tblLastAgreements.PayProcess = dbo.tblMiddlePay.KodMiddlePay
  LEFT JOIN dbo.tblAgreementKind
    ON dbo.tblAgreementKind.KodAgreementKind = tblLastAgreements.KodAgreementKind
  LEFT JOIN (
    SELECT DISTINCT NumR, Conditions,
      SUM(CASE WHEN NumPay = 0 AND SumP-Pay>0 AND PayRemark LIKE '%בונוס%' THEN SumP-pay ELSE 0 END) OVER(PARTITION BY NumR) AS SumBonus,
      SUM(CASE WHEN NumPay = 0 AND SumP-Pay>0  AND NOT PayRemark LIKE '%בונוס%' THEN SumP-pay ELSE 0 END) OVER(PARTITION BY NumR) AS SumSpecial,
      SUM(CASE WHEN NumPay > 0 AND SumP-Pay>0  AND DatePay IS NULL AND (GETDATE() > (CASE
     WHEN tblAgreementConditionAdvice.shotef = 1 THEN
       DateAdd(DAY, tblAgreementConditionAdvice.ashray+1,
               DateAdd(DAY,-1, Cast(CONVERT(VARCHAR(10),DATEPART(MM, DateAdd(month,1,tblAgreementConditionAdvicePay.DateP)),103)
                                    + '/' + '01/' + CONVERT(VARCHAR(10),DATEPART(YYYY, DateAdd(month,1,tblAgreementConditionAdvicePay.DateP)),103) AS DATETIME)))
     ELSE
       DateAdd(DAY, tblAgreementConditionAdvice.ashray+1,tblAgreementConditionAdvicePay.DateP)
     END)) THEN 1 ELSE 0 END) OVER(PARTITION BY NumR) AS NumOfDebt,
      SUM(CASE WHEN NumPay > 0 AND SumP-Pay>0  AND DatePay IS NULL AND (GETDATE() > (CASE
     WHEN tblAgreementConditionAdvice.shotef = 1 THEN
       DateAdd(DAY, tblAgreementConditionAdvice.ashray+1,
               DateAdd(DAY,-1, Cast(CONVERT(VARCHAR(10),DATEPART(MM, DateAdd(month,1,tblAgreementConditionAdvicePay.DateP)),103)
                                    + '/' + '01/' + CONVERT(VARCHAR(10),DATEPART(YYYY, DateAdd(month,1,tblAgreementConditionAdvicePay.DateP)),103) AS DATETIME)))
     ELSE
       DateAdd(DAY, tblAgreementConditionAdvice.ashray+1,tblAgreementConditionAdvicePay.DateP)
     END) ) THEN SumP ELSE 0 END) OVER(PARTITION BY NumR) AS Debt,
      STUFF((SELECT ', ' + tblRemarks.PayRemark AS [text()]FROM dbo.tblAgreementConditionAdvicePay tblRemarks WHERE tblRemarks.NumR = dbo.tblAgreementConditionAdvicePay.NumR AND tblRemarks.NumPay = 0 AND SumP-Pay>0 FOR XML PATH('')), 1, 1, '' ) AS Remark
    FROM dbo.tblAgreementConditionAdvicePay
    LEFT JOIN dbo.tblAgreementConditionAdvice
      ON dbo.tblAgreementConditionAdvice.IDagreement = dbo.tblAgreementConditionAdvicePay.NumR


     )AS tblSumAllPay
    ON tblLastAgreements.IDagreement = tblSumAllPay.NumR
  LEFT JOIN (
    SELECT KodCustomer, Text
    FROM tbltaxes
    ) tbltaxes
    ON
      dbo.tblCustomers.KodCustomer = tbltaxes.KodCustomer
WHERE CustomerStatus = 2
AND agreementFinish = 0
ORDER BY NameTeam'''
