﻿using Ionic.Zip;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace atlantis2webuntis
{
    public class Schuelers : List<Schueler>
    {
        internal void GeneriereImportdateiFürWebuntis(string datei)
        {
            Encoding enc = Encoding.GetEncoding("UTF-8");

            using (StreamWriter writer = new StreamWriter(datei, true, enc))
            {
                Console.Write(("Die Datei wird geschrieben: " + datei).PadRight(75, '.'));

                writer.WriteLine(@"Schlüssel;E-Mail;Familienname;Vorname;Klasse;Kurzname;Geschlecht;Geburtsdatum;Eintrittsdatum;Austrittsdatum;Telefon;Mobil;Strasse;PLZ;Ort;ErzName;ErzMobil;ErzTelefon;Volljährig;BetriebName;BetriebStrasse;BetriebPlz;BetriebOrt;BetriebTelefon;O365Identität;Benutzername");

                foreach (var schueler in this)
                {
                    writer.WriteLine(schueler.SchreibeDatensatz());
                }

                Console.WriteLine((" ok").PadLeft(30, '.'));
            }
        }

        internal void MoodlekurseAusKlassenAnlegen(string moodlekurseAusKlassenAnlegen)
        {
            Encoding enc = Encoding.GetEncoding("UTF-8");

            using (StreamWriter writer = new StreamWriter(moodlekurseAusKlassenAnlegen, true, enc))
            {
                Console.Write(("Die Datei wird geschrieben: " + moodlekurseAusKlassenAnlegen).PadRight(75, '.'));

                writer.WriteLine(@"shortname;fullname;summary;category");

                foreach (string klasse in (from s in this select s.Grade).Distinct().ToList())
                {
                    writer.WriteLine(klasse + ";" + klasse + ";Angelegt am " + DateTime.Now.ToShortDateString() + ";2");
                }

                Console.WriteLine((" ok").PadLeft(30, '.'));
            }
        }

        internal void SchülerInKlassenMoodlekurseEinschreiben(string schülerInKlassenMoodlekurseEinschreiben)
        {
            Encoding enc = Encoding.GetEncoding("UTF-8");

            using (StreamWriter writer = new StreamWriter(schülerInKlassenMoodlekurseEinschreiben, true, enc))
            {
                Console.Write(("Die Datei wird geschrieben: " + schülerInKlassenMoodlekurseEinschreiben).PadRight(75, '.'));

                writer.WriteLine(@"username;password;lastname;firstname;email;auth;course1;group1;department;city;deleted");

                foreach (var s in this)
                {
                    writer.WriteLine(s.Kurzname + ";123;" + s.Name + ";" + s.Firstname + ";" + s.Mail + ";ldap;" + s.Grade + ";" + s.Grade + ";" + s.Grade + ";" + DateTime.Now.ToString("yyyyMMdd") + ";" + (s.Austrittsdatum.Date == DateTime.Now.Date ? 1 : 0));
                }

                Console.WriteLine((" ok").PadLeft(30, '.'));
            }
        }


        internal void ZippeBilder(string pfadZuAtlantisFotos)
        {
            using (ZipFile zip = new ZipFile())
            {
                Console.Write("Vorhandene Bilder ".PadRight(75, '.'));

                UTF8Encoding utf8NoBom = new UTF8Encoding(false);
                var filePath = "import.csv";

                File.WriteAllText(filePath, "\"id\",\"custom_id\",\"email\",\"path\"" + Environment.NewLine, utf8NoBom);

                List<string> bereitsVerarbeiteteBilder = new List<string>();
                bereitsVerarbeiteteBilder = (Properties.Settings.Default.bereitsVerarbeiteteBilder).Split(',').ToList();

                int bilderHinzugefügt = 0;
                int bilderBereitsVorhandenUndNichtErneutHinzugefügt = 0;

                foreach (var schueler in this)
                {
                    var allFiles = Directory.GetFiles(pfadZuAtlantisFotos, "*" + schueler.Id + ".jpg", SearchOption.AllDirectories);

                    foreach (var item in allFiles)
                    {
                        if (schueler.Id.ToString().Length >= 6)
                        {
                            try
                            {
                                if ((from a in bereitsVerarbeiteteBilder where a == schueler.Id.ToString() select a).Count() == 0)
                                {
                                    zip.AddFile(item).FileName = schueler.Id + ".jpg";
                                    bilderHinzugefügt++;
                                    bereitsVerarbeiteteBilder.Add(schueler.Id.ToString());
                                    //Console.WriteLine("Bild hinzugefügt für " + schueler.Kurzname);

                                    if (schueler.Mail != null && schueler.Mail != "")
                                    {
                                        File.AppendAllText(filePath, "\"\",\"\",\"" + schueler.Mail + "\",\"" + schueler.Id + ".jpg" + "\"" + Environment.NewLine, utf8NoBom);
                                    }
                                }
                                else
                                {
                                    bilderBereitsVorhandenUndNichtErneutHinzugefügt++;
                                    // Console.WriteLine("bereits vorhanden");
                                }
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }
                }
                string f = "";

                foreach (var item in bereitsVerarbeiteteBilder)
                {
                    f = f + item + ",";
                }

                Properties.Settings.Default.bereitsVerarbeiteteBilder = f;
                Properties.Settings.Default.Save();

                Console.WriteLine((" " + bilderBereitsVorhandenUndNichtErneutHinzugefügt).PadLeft(30, '.'));

                Console.Write("Neu eingelesene und gezippte Bilder: ".PadRight(75, '.'));

                zip.Save("SchülerBilder-" + DateTime.Now.ToFileTime() + ".zip");

                zip.AddFile(filePath);

                string datei = "Geevoo-Import-" + DateTime.Now.ToFileTime() + ".zip";

                zip.Save(datei);

                Console.WriteLine((" " + bilderHinzugefügt).PadLeft(30, '.'));
            }
        }

        public static string SafeGetString(OleDbDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }

        public Schuelers(string connectionStringAtlantis, Betriebe betriebe, Adressen adressen, Klasses klasses, Bemerkungen bemerkungen)
        {
            using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
            {
                DataSet dataSet = new DataSet();

                OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"SELECT   
            schueler.pu_id AS AtlantisSchuelerId,
            schueler.gebname AS Geburtsname,
            schueler.gebort_lkrs AS Geburtsort,
            schueler.s_geburts_land AS Geburtsland,
            schue_sj.vorgang_akt_satz_jn AS AktuellJN,
schue_sj.s_schulabschluss_bos AS Versetzung1,
schue_sj.dat_versetzung AS VersetzungsDatumInDiesenBildungsgang,
schue_sj.s_austritts_grund_bdlg AS BBB,
schue_sc.dat_austritt AS Austrittsdatum,
schue_sc.s_ausb_ziel_erreicht AS BeruflicherAbschluss,
schue_sc.s_austritt_schulabschluss AS AllgemeinerAbschluss,
schue_sc.s_hoechst_schulabschluss AS LSHoechAllAb,
schueler.s_schwerstbehindert AS XX,
schue_sj.kl_id AS AtlantisKlasseId,			
schue_sj.fko_id AS AtlantisStundentafelId,			
schue_sj.durchschnitts_note_jz AS DurchschnittsnoteJahreszeugnis,			
			schueler.name_1 AS Nachname,
			schueler.name_2 AS Vorname,
			check_null (hole_schuljahr_rech ('', -1))												as schuljahr_vorher, 
			check_null (hole_schuljahr_rech ('', 0))												as Bezugsjahr,				/*  1 aktuelles SJ des Users */
			check_null (sv_km_kuerzel_wert ('s_typ_vorgang', schue_sj.s_typ_vorgang))	as Status,  				/*  2 */
			/* lfd_nr im cf_ realisiert */  																							/*  3 */ 
			(if check_null (klasse.klasse_statistik_name) <> ''	THEN
				 check_null (klasse.klasse_statistik_name)         ELSE
				 check_null (klasse.klasse)								ENDIF)				as Klasse,  					/*  4 */

         check_null (sv_km_kuerzel_wert ('s_berufs_nr_gliederung', 
                                 schue_sj.s_berufs_nr_gliederung))   				as Gliederung,  				/*  5 */   
         check_null (substr(schue_sj.s_berufs_nr,4,5))				  					as Fachklasse,  				/*  6 Stelle 4-8 */  
			''																								as Klassenart,					/*  7 nicht BK */
			check_null (sv_km_kuerzel_wert ('s_klasse_art', klasse.s_klasse_art))	as OrgForm,  					/*  8 */
         check_null (sv_km_kuerzel_wert ('jahrgang', schue_sj.s_jahrgang))			as AktJahrgang,  				/*  9 */   
			check_null (sv_km_kuerzel_wert ('s_art_foerderungsbedarf', schueler.s_art_foerderungsbedarf))
																											as Foerderschwerp,			/* 10 */ 
			(if check_null (schueler.s_schwerstbehindert) = ''	THEN
				 '0'                 									ELSE
				 '1'															ENDIF)					as Schwerstbeh,				/* 11 */ 
			''																								as Reformpdg,					/* 12 */ 
			(if sv_steuerung ('s_unter', schueler.s_unter) = '$JVA' THEN
				 '1'                 									ELSE
				 '0'															ENDIF)		 				as JVA,							/* 13 */ 
         check_null_10(adresse.plz )                                             	as Plz, 				 			/* 14 */ 
			hole_schueler_ort_bundesland (adresse.s_gem_kz, adresse.ort, adresse.lkz)	as Ort,							/* 15 */
         schueler.dat_geburt			                    										as Gebdat,   					/* 16 */
         check_null (sv_km_kuerzel_wert ('s_geschl' , schueler.s_geschl))				as Geschlecht,   				/* 17 */   
         check_null (sv_km_kuerzel_wert ('s_staat'  , schueler.s_staat))				as Staatsang,   				/* 18 */
         
		   check_null (sv_km_kuerzel_wert ('s_bekennt', schueler.s_bekennt))       as Religion,  					/* 19 */
			schue_sj.dat_rel_anmeld																	as Relianmeldung,				/* 20 */
			schue_sj.dat_rel_abmeld																	as Reliabmeldung,				/* 21 */

			(if Aufnahmedatum_Bildungsgang is null THEN     		
				 Aufnahmedatum_Schule					ELSE
				 Aufnahmedatum_Bildungsgang			ENDIF)        							as	Aufnahmedatum,   			/* 22 */  

			(select max (lehr_sc.ls_kuerzel)
            from lehr_sc, kl_ls, lehrer
        	  where lehr_sc.ls_id 		= kl_ls.ls_id                                                   
        		 and kl_ls.kl_id 			= klasse.kl_id
				 and kl_ls.s_typ_kl_ls 	= '0'
             and lehrer.le_id 		= lehr_sc.le_id) 										as Labk,							/* 23 */
						

			(select adresse.plz
			   from adresse, betrieb, pj_bt  
			  where adresse.ad_id 		= betrieb.id_hauptadresse  
				 and pj_bt.bt_id 			= betrieb.bt_id    
				 and pj_bt.s_typ_pj_bt 	= '0'       
				 and pj_bt.pj_id 			= schue_sj.pj_id)										as ausbildort,  				/* 24 */                       
				
			hole_schueler_betriebsort (schue_sj.pj_id, 'ort')								as betriebsort,  				/* 25 */                       
				
			/* Kapitel der zuletzt besuchten Schule */
         check_null (sv_km_kuerzel_wert ('s_herkunfts_schule', 
                                 schue_sj.s_herkunfts_schule))                   as	LSSchulform,/* 26 */  
			left (schue_sj.vo_s_letzte_schule, 6) 						      		  		as LSSchulnummer,				/* 27 */
			check_null (sv_km_kuerzel_wert ('s_berufl_vorbildung_glied', 
                                 schue_sj.s_berufl_vorbildung_glied))        		as	LSGliederung,   			/* 28 */  
			substr (schue_sj.s_berufl_vorbildung, 4, 5) 			   						as	LSFachklasse,   			/* 29 */  

			''																								as LSKlassenart,				/* 30 */
			''																								as LSReformpdg,				/* 31 */
			Null																							as LSSchulentl,				/* 32 */		

			check_null (sv_km_kuerzel_wert ('s_abgang_jg', schue_sj.vo_s_jahrgang))	
																        									as	LSJahrgang,   				/* 33 */  

			(if Gliederung || Fachklasse = VOGliederung || VOFachklasse then 
             (if AktJahrgang = vojahrgang and Schueler_le_schuljahr_da = 'J' then 'W' else 'V' endif)   else
		      right (check_null (sv_km_kuerzel_wert ('s_schulabschluss', schue_sc.s_hoechst_schulabschluss)), 1) ENDIF)
																             							as LSQual,    					/* 34 */

			'0'																							as LSVersetz,					/* 35 */ 
	 
			/* Kapitel für das abgelaufene Schuljahr */
			check_null ((if (Fall_Bezugsjahr = '1')				THEN
				(SELECT klasse.klasse									/* aus Vorjahr lesen */  
					FROM klasse, schue_sj  
				  WHERE schue_sj.kl_id = klasse.kl_id     
					 and schue_sj.pj_id = VOpj_id)					ELSE
				/* Fall_schuljahr_vorher */
				klasse.klasse												ENDIF))				 	as VOKlasse,    				/* 36 */

			check_null ((if (Fall_Bezugsjahr = '1') 				THEN
				(select schue_sj.s_berufs_nr_gliederung			/* aus Vorjahr lesen */
					from schue_sj
				  where schue_sj.pj_id 	= VOpj_id)					ELSE
				/* Fall_schuljahr_vorher */
				schue_sj.s_berufs_nr_gliederung						ENDIF)) 					as VOGliederung,  			/* 37 */

			check_null ((if (Fall_Bezugsjahr = '1') 				THEN
				(select substr (schue_sj.s_berufs_nr, 4, 5)		/* aus Vorjahr lesen */ 
					from schue_sj
				  where schue_sj.pj_id 	= VOpj_id)					ELSE
				/* Fall_schuljahr_vorher */
				substr (schue_sj.s_berufs_nr, 4, 5)					ENDIF)) 					as VOFachklasse, 				/* 38 */

			check_null ((if (Fall_Bezugsjahr = '1')				THEN
				(SELECT sv_km_kuerzel_wert ('s_klasse_art', klasse.s_klasse_art)/* aus Vorjahr lesen */  
					FROM klasse, schue_sj  
				  WHERE schue_sj.kl_id = klasse.kl_id     
					 and schue_sj.pj_id = VOpj_id)					ELSE
				/* Fall_schuljahr_vorher */
				sv_km_kuerzel_wert ('s_klasse_art', klasse.s_klasse_art) ENDIF))		as VOOrgForm,    				/* 39 */
			''																								as VOKlassenart,				/* 40 nicht BK*/

			check_null ((if (Fall_Bezugsjahr = '1')				THEN
				(SELECT sv_km_kuerzel_wert ('jahrgang', schue_sj.s_jahrgang)/* aus Vorjahr lesen */  
					FROM klasse, schue_sj  
				  WHERE schue_sj.kl_id = klasse.kl_id     
					 and schue_sj.pj_id = VOpj_id)					ELSE
				/* Fall_schuljahr_vorher */
				sv_km_kuerzel_wert ('jahrgang', schue_sj.s_jahrgang) ENDIF))		   as VOJahrgang,    			/* 41 */

			check_null (sv_km_kuerzel_wert ('s_art_foerderungsbedarf_vj', 
                                 schueler.s_art_foerderungsbedarf_vj))   			as VOFoerderschwerp,			/* 42 */ 
			'0'																							as VOSchwerstbeh,				/* 43 */ 
			''																								as VOReformpdg,				/* 44 nicht BK*/
				
			hole_schueler_bildungsgang (schueler.pu_id, klasse.s_bildungsgang, schue_sj.dat_austritt, 'austritt')
																											as EntlDatum,					/* 45 */				

			check_null (sv_km_kuerzel_wert ('s_austritts_grund_bdlg', 
                                 schue_sj.s_austritts_grund_bdlg))               as Zeugnis_JZ, 				/* 46 */

			''																								as Schulpflichterf,			/* 47 nicht BK */ 			

			''																								as Schulwechselform,			/* 48 nicht BK */ 		

			''																								as Versetzung,					/* 49 */ 		

			(IF s_geburts_land is not null AND s_geburts_land <> '000' THEN
             string (year (dat_zuzug))
          ELSE
             ''
          ENDIF
			)																								as JahrZuzug_nicht_in_D_geboren,	/* 50 */

			string (year (dat_eintritt_gs))														as JahrEinschulung,					/* 51 */

			''																								as JahrWechselSekI,					/* 52 */

			(IF (s_geburts_land is not null AND s_geburts_land <> '000') OR (s_geburts_land is null AND dat_zuzug is not null) THEN
             '1'
          ELSE
             '0'
          ENDIF
			)																							   as Zugezogen,							/* 53 */


			/* Hilfsfelder */
			(SELECT max (adr_mutter.ad_id) 
            FROM adresse adr_mutter 
           WHERE adr_mutter.pu_id		= schueler.pu_id
             AND adr_mutter.s_typ_adr	= 'M'
			)																								as ad_id_mutter,


(select max (adresse.strasse)
			   from adresse  
			  where adresse.pu_id = schueler.pu_id)										as strasse, 

(select max (adresse.tel_1)
			   from adresse  
			  where adresse.pu_id = schueler.pu_id)										as telefon, 

			(SELECT adr_mutter.s_herkunftsland_adr 
            FROM adresse adr_mutter 
           WHERE adr_mutter.ad_id		= ad_id_mutter
			)																								as herkunftsland_mutter,			/* 54 */

			(SELECT max (adr_vater.ad_id) 
            FROM adresse adr_vater
           WHERE adr_vater.pu_id			= schueler.pu_id
             AND adr_vater.s_typ_adr	= 'V'
			)																								as ad_id_vater,

			(SELECT adr_vater.s_herkunftsland_adr 
            FROM adresse adr_vater
           WHERE adr_vater.ad_id			= ad_id_vater
			)																								as herkunftsland_vater,				/* 55 */

			(IF herkunftsland_mutter is not null AND herkunftsland_mutter <> '000'
             OR
				 herkunftsland_vater  is not null AND herkunftsland_vater  <> '000' THEN
             '1'
          ELSE
             '0'
          ENDIF
         )																								as Elternteilzugezogen,				/* 56 */

			(IF schueler.s_muttersprache is not null AND 
             schueler.s_muttersprache <> 'DE'     AND 
             schueler.s_muttersprache <> '000' THEN
             schueler.s_muttersprache
          ELSE
             ''
          ENDIF
         )																								as Verkehrssprache,					/* 57 */

			''																								as Einschulungsart					/* 58 */,
			''																								as Grundschulempfehlung				/* 59 */,

			schue_sj.pj_id																				as pj_id,
			check_null (sv_km_kuerzel_wert ('s_religions_unterricht', 
                                 schue_sj.s_religions_unterricht))					as s_religions_unterricht,
     		schue_sc.dat_eintritt	                   										as	Aufnahmedatum_schule,   			

			hole_schueler_bildungsgang (schueler.pu_id, klasse.s_bildungsgang, schue_sj.dat_eintritt, 'eintritt')
							     			                   										as	Aufnahmedatum_Bildungsgang,

			(SELECT max (pj_id) 
			   FROM schue_sj  
			  WHERE schue_sj.pu_id 					= schueler.pu_id 		and
					  schue_sj.s_typ_vorgang 		IN ('A', 'G', 'S')	and
					  schue_sj.vorgang_schuljahr 	= schuljahr_vorher) 						as VOpj_id,

			test_austritt (schue_sj.dat_austritt)												as ausgetreten,
		
			check_null ((SELECT noten_kopf.s_bestehen_absprf || noten_kopf.s_zusatzabschluss
			   FROM noten_kopf, schue_sj  
			  WHERE noten_kopf.s_typ_nok 			= 'HZ'
				 AND schue_sj.pj_id 					= noten_kopf.pj_id
				 AND schue_sj.pj_id 					= VOpj_id))            					as Zeugnis_HZ,


			check_null ((if ( schue_sj.vorgang_schuljahr 	= check_null (hole_schuljahr_rech ('',  0)) /* '2005/06' */
	  				AND schue_sj.vorgang_akt_satz_jn = 'J'
	  				AND ausgetreten 						= 'N') 		THEN
				 '1'                 									ELSE
				 '0'															ENDIF))					as Fall_Bezugsjahr,			

			check_null ((if ( schue_sj.vorgang_schuljahr 	= check_null (hole_schuljahr_rech ('', -1)) /* '2004/05' */
				   AND schue_sj.s_typ_vorgang 	IN ('A', 'G', 'S')
				   AND ausgetreten 					= 'J') 			THEN
				 '1'                 									ELSE
				 '0'															ENDIF))					as Fall_schuljahr_vorher,
			
			check_null ((if ( VOJahrgang <> '')							 			THEN
				 'J'                 									ELSE
				 'N'															ENDIF))					as Schueler_le_schuljahr_da,	
			
			check_null ((if ( Fachklasse <> VOFachklasse
					OR Gliederung <> VOGliederung)		 			THEN
				 'J'                 									ELSE
				 'N'															ENDIF))					as Schueler_Berufswechsel,	
			
			check_null ((if ( VoGliederung = 'C05' 
				  AND Gliederung = 'C06')		 						THEN
				 'J'                 									ELSE
				 'N'															ENDIF))					as Vorjahr_C05_Aktjahr_C06,		
			
			check_null (sv_km_kuerzel_wert ('jahrgang', schue_sj.s_jahrgang))				as schueler_jahrgang, 

			check_null ((if (Fall_Bezugsjahr = '1') 				THEN
				(select schue_sj.s_jahrgang			/* aus Vorjahr lesen */
					from schue_sj
				  where schue_sj.pj_id 	= VOpj_id)					ELSE
				/* Fall_schuljahr_vorher */
				schue_sj.s_jahrgang										ENDIF)) 					as VOSchueler_Jahrgang,

  			(IF EXISTS (SELECT 1 FROM schue_sj_info  
                      WHERE schue_sj_info.info_gruppe = 'MASSNA'
						      AND schue_sj_info.pj_id = schue_sj.pj_id) THEN '1' ELSE '0' ENDIF) as Massnahmetraeger,	/* 60 */

         check_null (sv_km_kuerzel_wert ('s_geschl', schueler.s_betreuungsart))	as Betreuung,   						/* 61 */   

  			(IF EXISTS (SELECT 1 FROM schueler_info  
                      WHERE schueler_info.info_gruppe = 'PUBEM'
                        AND schueler_info.s_typ_puin  = 'BKAZVO'
								AND schueler_info.betrag IN (0, 6, 12, 18)
						      AND schueler_info.pu_id = schueler.pu_id) THEN 
		   (SELECT konv_dec_string(schueler_info.betrag, 0)
			   FROM schueler_info  
			  WHERE schueler_info.puin_id = (SELECT max(puin.puin_id)  
														  FROM schueler_info puin 
														 WHERE puin.info_gruppe = 'PUBEM' 
														   AND puin.s_typ_puin = 'BKAZVO'
														   AND puin.betrag IN (0, 6, 12, 18)
														   AND puin.pu_id  = schueler.pu_id)) ELSE '0' ENDIF)	as BKAZVO,	/* 62 */   
			check_null (sv_km_kuerzel_wert ('s_art_foerderungsbedarf2', 
                                 schueler.s_art_foerderungsbedarf2))  	as Foerderschwerp2,						/* 63 */   
			check_null (sv_km_kuerzel_wert ('s_art_foerderungsbedarf_vj2',
                                 schueler.s_art_foerderungsbedarf_vj2)) as VOFoerderschwerp2,					/* 64 */   
			(IF schue_sc.berufsabschluss_jn = 'J' THEN 'Y' ELSE '' ENDIF)	as Berufsabschluss, 						/* 65 */   

			'Atlantis' 																		as Produktname, 						   /* 66 */   
			(SELECT 'V' || version.version_nr  
			   FROM version
           WHERE version.datum  = (SELECT max(v.datum) FROM version v)) as Produktversion,     			      /* 67 */   

         check_null (sv_km_kuerzel_wert ('s_bereich', klasse.s_bereich)) as Adressmerkmal,                  /* 68 */
			(if sv_steuerung ('s_unter', schueler.s_unter) = '$IN' THEN
				 '1'                 									ELSE
				 ''															ENDIF)		 as Internat,								/* 69 */ 
         '0'																			    as Koopklasse                      /* 70 */

  FROM   schueler,   
         schue_sc,  
         schue_sj,   
			schule,
         klasse,   
         adresse

 WHERE   schue_sj.kl_id 					= klasse.kl_id    			     
     AND schue_sc.sc_id 					= schule.sc_id  			    
     AND schue_sc.ps_id 					= schue_sj.ps_id  			     
     AND schue_sc.pu_id 					= schueler.pu_id  			     
	  AND adresse.ad_id 						= schueler.id_hauptadresse  
     AND (
			 (schue_sj.vorgang_schuljahr 	= check_null (hole_schuljahr_rech ('',  0)) /* aktuelles Jahr */
	  
			 )
          OR
          (
			  schue_sj.vorgang_schuljahr 	= check_null (hole_schuljahr_rech ('', -1)) /* letzte jahr */
	  AND   schue_sj.s_typ_vorgang 			IN ('G', 'S')
	  AND   ausgetreten = 'J'
			 )
         ) 
ORDER BY ausgetreten DESC, klasse, schueler.name_1, schueler.name_2", connection);

                Console.Write("Schüler*innen aus Atlantis ".PadRight(75, '.'));

                connection.Open();
                schuelerAdapter.Fill(dataSet, "DBA.schueler");

                foreach (DataRow theRow in dataSet.Tables["DBA.schueler"].Rows)
                {
                    string vorname = theRow["Vorname"] == null ? "" : theRow["Vorname"].ToString();
                    string nachname = theRow["Nachname"] == null ? "" : theRow["Nachname"].ToString();

                    if (vorname.Length > 1 && nachname.Length > 1)
                    {
                        var schueler = new Schueler();
                        schueler.Id = theRow["AtlantisSchuelerId"] == null ? -99 : Convert.ToInt32(theRow["AtlantisSchuelerId"]);
                        schueler.Name = theRow["Nachname"] == null ? "" : theRow["Nachname"].ToString();
                        schueler.Firstname = theRow["Vorname"] == null ? "" : theRow["Vorname"].ToString();
                        schueler.Grade = theRow["Klasse"] == null ? "" : theRow["Klasse"].ToString();
                        schueler.Birthday = theRow["Gebdat"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Gebdat"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        schueler.Telefon = theRow["telefon"] == null ? "" : theRow["telefon"].ToString();
                        schueler.Kurzname = schueler.generateKurzname();
                        schueler.Mail = schueler.Kurzname + "@students.berufskolleg-borken.de";
                        schueler.Eintrittsdatum = theRow["Aufnahmedatum"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Aufnahmedatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                        schueler.AktuellJN = theRow["AktuellJN"] == null ? "" : theRow["AktuellJN"].ToString();

                        schueler.Austrittsdatum = theRow["Austrittsdatum"].ToString().Length < 3 ? new DateTime((DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 : DateTime.Now.Year), 7, 31) : DateTime.ParseExact(theRow["Austrittsdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        schueler.Volljährig = schueler.Birthday.AddYears(18) > DateTime.Now ? false : true;
                        var geschlecht345 = theRow["Geschlecht"] == null ? "" : theRow["Geschlecht"].ToString();
                        schueler.Geschlecht = geschlecht345 == "3" ? "m" : geschlecht345 == "4" ? "w" : "d";
                        schueler.Relianmeldung = theRow["Relianmeldung"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Relianmeldung"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        schueler.Reliabmeldung = theRow["Reliabmeldung"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Reliabmeldung"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                        schueler.Status = theRow["Status"] == null ? "" : theRow["Status"].ToString();
                        schueler.Bezugsjahr = Convert.ToInt32(theRow["Bezugsjahr"].ToString().Substring(0, 4)) - Convert.ToInt32(theRow["Fall_schuljahr_vorher"]);

                        schueler.Betrieb = (from b in betriebe where b.SchuelerIdAtlantis == schueler.Id select b).FirstOrDefault();

                        schueler.Adressen = new List<Adresse>();

                        schueler.Adressen.AddRange(from a in adressen where a.IdAtlantis == schueler.Id select a);

                        if (schueler.Name == "Röhricht")
                        {
                            string a = "";
                        }

                        // Beurlaubte Schüler werden nicht nach Untis exportiert

                        bool beurlaubt = (from b in bemerkungen where b.AtlantisSchuelerId == schueler.Id where b.Kürzel == "B" select b).Any();

                        if (!beurlaubt && schueler.Bezugsjahr == (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1) && schueler.Status != "VB" && schueler.Status != "8" && schueler.Status != "9" && schueler.Grade != "Z" && schueler.AktuellJN == "J")
                        {
                            // Duplikate werden verhindert.

                            if (!(from s in this where s.Id == schueler.Id select s).Any())
                            {
                                this.Add(schueler);
                            }
                        }
                    }
                }
                Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));

                // Duplikate finden.

                foreach (var item in this.GroupBy(x => x.Id).Where(g => g.Count() > 1).Select(y => y.Key).ToList())
                {
                    var s = (from i in this where item == i.Id select i).ToList();

                    foreach (var ss in s)
                    {
                        Console.WriteLine("Duplikat: Der SuS kommt mehrfach vor:" + ss.Id + " " + ss.Kurzname + " " + ss.Grade + " " + ss.Eintrittsdatum.ToShortDateString() + " " + ss.AktuellJN);
                    }
                }
            }
        }

        public Schuelers()
        {
            var user = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1];

            var datei = (from f in Directory.GetFiles(@"c:\users\" + user + @"\Downloads", "*.csv", SearchOption.AllDirectories) where f.ToLower().Contains("student") orderby File.GetLastWriteTime(f) select f).LastOrDefault();

            //string datei = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Student.csv";

            if (!File.Exists(datei))
            {
                Console.WriteLine("Die Datei " + datei + " existiert nicht.");
                Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
                Console.WriteLine("1. sich als admin anmelden");
                Console.WriteLine("2. auf Stammdaten > Schüler*innen klicken");
                Console.WriteLine("3. 'Zum Kalenderdatum' anhaken");
                Console.WriteLine("4. Berichte Schüler als CSV im Download-Ordner speichern");
                Console.WriteLine("ENTER beendet das Programm.");
                Console.ReadKey();  
                Environment.Exit(0);
            }
            else
            {
                if (System.IO.File.GetLastWriteTime(datei).Date != DateTime.Now.Date)
                {
                    Console.WriteLine("Die Datei " + datei + " existiert nicht.");
                    Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
                    Console.WriteLine("1. sich als admin anmelden");
                    Console.WriteLine("2. auf Stammdaten > Schüler*innen klicken");
                    Console.WriteLine("3. 'Zum Kalenderdatum' anhaken");
                    Console.WriteLine("4. Berichte Schüler als CSV exportieren nach " + datei);
                    Console.WriteLine("ENTER beendet das Programm.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }
            }

            if (datei != null)
            {
                Console.WriteLine("Ausgewertete Datei: " + (Path.GetFileName(datei) + " ").PadRight(38, '.') + ". Erstell-/Bearbeitungszeitpunkt heute um " + System.IO.File.GetLastWriteTime(datei).ToShortTimeString());
            }

            using (StreamReader reader = new StreamReader(datei))
            {
                Console.Write("Schüler*innen aus Webuntis ".PadRight(75, '.'));

                while (true)
                {
                    string line = reader.ReadLine();

                    if (line == null)
                    {
                        break;
                    }

                    try
                    {
                        Schueler schueler = new Schueler();
                        var x = line.Split('\t');
                        schueler.Kurzname = x[0];
                        schueler.Name = x[1];
                        schueler.Firstname = x[2];
                        schueler.Geschlecht = x[3];
                        schueler.Geburtsdatum = x[4];
                        schueler.Grade = x[5];
                        schueler.Eintrittsdatum = x[6] == "" ? new DateTime() : DateTime.ParseExact(x[6], "dd.MM.yyyy", CultureInfo.InvariantCulture);
                        schueler.Austrittsdatum = x[7] == "" ? new DateTime() : DateTime.ParseExact(x[7], "dd.MM.yyyy", CultureInfo.InvariantCulture);
                        schueler.Id = Convert.ToInt32(x[10]);
                        schueler.Telefon = "";
                        schueler.Strasse = "";
                        schueler.Plz = "";
                        schueler.Ort = "";
                        schueler.Volljährig = false;
                        this.Add(schueler);
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));
        }
    }
}