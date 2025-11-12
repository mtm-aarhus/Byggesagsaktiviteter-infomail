select
	TRIM(sager1.Sagsnummer) 'Sagsnummer'
	,CAST(sager1.Sagsdato AS date) 'Sagsdato'
	,CAST(info1.Afgoerelsedato AS date) 'Afgørelsesdato'
	,CAST(info1.AfslutningsDato AS date) 'Afslutningsdato'
	,TRIM(info1.Sagsadresse) 'Sagsadresse'
	,TRIM(sager1.Titel) 'Sagsoverskrift'
	,TRIM(info1.Sagsart) 'Sagsart'
	,TRIM(sager1.Sagsbehandler_brugernavn) 'Sagsbehandler'
	,TRIM(Tilstand) 'Tilstand'
	
FROM 
	Service_NOVA2.LIS_Byg_sager sager1
	INNER JOIN Service_NOVA2.LIS_Byg_info info1 on sager1.SagsUUID = info1.SagsUUID
	
	where sager1.SagsUUID in (

		select sager2.SagsUUID
		from 
			Service_NOVA2.LIS_Byg_sager sager2
			inner join Service_NOVA2.LIS_Byg_info info2 ON sager2.SagsUUID = info2.SagsUUID
			left join Service_NOVA2.LIS_Byg_Opgaver opgaver on sager2.SagsUUID = opgaver.SagsUUID
		where
			sager2.Tilstand not in ('Afsluttet')
			AND info2.AfslutningsDato is null
			AND sager2.Sagsnummer NOT LIKE '%skabelon%'
			AND sager2.Sagsnummer NOT LIKE '%S2023-24573%'
			AND sager2.Titel NOT LIKE '%Meddelelser om arrangementer 2023%'
			AND sager2.Titel NOT LIKE 'TESTSAG%'
			AND sager2.Titel NOT LIKE 'TEST -%'
			AND sager2.Titel NOT LIKE 'TEST %'
			AND sager2.Titel NOT LIKE 'TEST,%'
			AND sager2.Titel NOT LIKE 'Lokalplan X%'
			AND sager2.Titel NOT LIKE '%Husnummerering%'
			AND sager2.Titel NOT LIKE '%filarkiv%'
			AND sager2.Titel NOT LIKE '%aktindsigt%'
			AND sager2.Titel not in ('Kims (KMD) testsag','(TEST KMD) test kmd','KMD test (Katrine)','KMD Test (Kim)')
			AND sager2.Titel NOT LIKE 'BBR%'
			and (opgaver.Afsluttetdato is null and opgaver.Fristdato is null)
			AND sager2.Sagsbehandler_brugernavn <> 'Hanne Eline Povlsen'

			AND sager2.SagsUUID IN (
				SELECT sager4.SagsUUID
				FROM LOIS.Service_NOVA2.LIS_Byg_sager sager4
				INNER JOIN LOIS.Service_NOVA2.LIS_Byg_Opgaver opgaver3 
					ON sager4.SagsUUID = opgaver3.SagsUUID
				WHERE opgaver3.[Status] = 'Igang' OR opgaver3.[Status] = 'I gang'
				GROUP BY sager4.SagsUUID
				HAVING COUNT(*) = SUM(CASE WHEN opgaver3.Titel LIKE '%Tidsreg%' THEN 1 ELSE 0 END)
			)
	)

order by Sagsnummer