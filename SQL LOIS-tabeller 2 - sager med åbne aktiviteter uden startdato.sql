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
			-- frav�lg sager med status "afsluttet"
			sager2.Tilstand not in ('Afsluttet')

			-- frav�lg sager hvor afslutningsdato er sat
			AND info2.AfslutningsDato is null

			-- frav�lg sager hvor ordet "skabelon" indg�r i sagsnummeret
			AND sager2.Sagsnummer NOT LIKE '%skabelon%'

			-- frav�lg sager hvor titlen indikerer at det er en testsag
			AND sager2.Titel NOT LIKE 'TESTSAG%'
			AND sager2.Titel NOT LIKE 'TEST -%'
			AND sager2.Titel NOT LIKE 'TEST %'
			AND sager2.Titel NOT LIKE 'TEST,%'

			-- frav�lg sager hvor titlen indeholder "lokalplan x"
			AND sager2.Titel NOT LIKE 'Lokalplan X%'

			-- frav�lg sager hvor titlen indeholder ordet "husnummerering"
			AND sager2.Titel NOT LIKE '%Husnummerering%'

			-- frav�lg sager hvor titlen indeholder ordet "filarkiv"
			AND sager2.Titel NOT LIKE '%filarkiv%'

			-- frav�lg sager hvor titlen indeholder ordet "aktindsigt"
			AND sager2.Titel NOT LIKE '%aktindsigt%'

			-- frav�lg KMD testsager
			AND sager2.Titel not in ('Kims (KMD) testsag','(TEST KMD) test kmd','KMD test (Katrine)','KMD Test (Kim)')

			-- frav�lg alle sager der starter med "BBR"
			AND sager2.Titel NOT LIKE 'BBR%'

			-- frav�lg opgaver hvor afslutningsdato og fristdato er sat
			and (opgaver.Afsluttetdato is null and opgaver.Fristdato is null)

			-- frav�lg sagsbehandler Hanne Eline Povlsen
			AND sager2.Sagsbehandler_brugernavn <> 'Hanne Eline Povlsen'

			-- frav�lg BBR-sager i sagsart
			AND info2.Sagsart NOT LIKE 'BBR%'

			-- frav�lg opgaver hvor startdato er sat og hvor afslutningsdato samtidig ikke er sat
			and sager2.SagsUUID not in (
				select
					sager3.SagsUUID
				from
					Service_NOVA2.LIS_Byg_sager sager3
					left join Service_NOVA2.LIS_Byg_Opgaver opgaver2 on sager3.SagsUUID = opgaver2.SagsUUID
				where
					opgaver2.Startdato is not null
					and opgaver2.Afsluttetdato is null
			)

			-- frav�lg opgaver hvor opgavens status er "S" og hvor opgavens titel samtidig ikke starter med "tidsreg"
			and sager2.SagsUUID NOT IN (
				Select sager4.SagsUUID
				FROM
					LOIS.Service_NOVA2.LIS_Byg_sager sager4
					INNER JOIN LOIS.Service_NOVA2.LIS_Byg_Opgaver opgaver3 On sager4.SagsUUID = opgaver3.SagsUUID
				WHERE
					opgaver3.[Status] = 'S'
					AND opgaver3.Titel NOT LIKE 'Tidsreg%'
			)
	)

order by Sagsnummer

