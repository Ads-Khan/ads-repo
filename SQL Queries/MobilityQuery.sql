

-- Uses LcsCDR

SELECT        
      [RegisterTime]
      ,[UserUri]
      ,[UserUriType]
      ,[ClientVersion]
      ,[ClientType]
      ,[ClientCategory]
      ,[Registrar]
	  ,[EdgeServer]
      ,[Pool]
      ,[IsInternal]
	  ,(Case 
	  WHEN [IsInternal] = 1 THEN 'Yes'
	  WHEN [IsInternal] = 0 THEN 'No'
	  end) AS 'User Inside?'
   
  FROM [LcsCDR].[dbo].[RegistrationView]

  WHERE [IsInternal] = 0
  --WHERE [ClientType] BETWEEN 16399 AND 16402

  ORDER BY [RegisterTime] DESC