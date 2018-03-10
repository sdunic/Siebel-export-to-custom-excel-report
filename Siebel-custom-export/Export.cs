using System;
using System.IO;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace Siebel_custom_report
{
    public partial class Export : Form
    {
        public Export()
        {
            InitializeComponent();
        }

        private string GetExcelColumnName(int broj)
        {
            int dividend = broj;
            string name = "";
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                name = Convert.ToChar(65 + modulo).ToString() + name;
                dividend = (int)((dividend - modulo) / 26);
            }
            return name;
        }

        bool splitBillerActive = false;
        int sbBroj = 0;

        private void btnReport_Click(object sender, EventArgs e)
        {
            if (String.Compare(btnExport.Text, "Zatvori") == 0) { Application.Exit(); }
            else
            {
                loadInfo.Text = "Učitavam podatke";
                Cursor = Cursors.WaitCursor;
                List<Siebel_export> dataSiebelExport = new List<Siebel_export>();
                dataSiebelExport.Clear();

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excelPath.ToString());
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;


                Excel.Range a1 = xlWorksheet.Cells[1, 4];
                Excel.Range a2 = xlWorksheet.Cells[rowCount, 4];
                xlWorksheet.get_Range(a1, a2).NumberFormat = "#";

                object[,] _2dArray = new object[rowCount - 1, colCount - 1];
                _2dArray = xlRange.Value2;

                

                Siebel_export row_siebel_export = new Siebel_export();
                List<Proizvodi> unique_proizvodi = new List<Proizvodi>();
                List<string> unique_sims = new List<string>();
                List<string> headers = new List<string>();

                for (int k = 1; k <= colCount; k++)
                {
                    headers.Add(_2dArray[1, k].ToString());
                }
                int kor_naplata = headers.IndexOf("Korisnik za naplatu") + 1;
                int kor_usluga = headers.IndexOf("Korisnik za uslugu") + 1;
                int broj_telefona = headers.IndexOf("Broj telefona") + 1;
                int dat_akt = headers.IndexOf("Datum aktivacije") + 1;
                int dat_deakt = headers.IndexOf("Datum deaktivacije") + 1;
                int status = headers.IndexOf("Status") + 1;
                int proizv = headers.IndexOf("Proizvod") + 1;
                int prof_napl = headers.IndexOf("Profil naplate") + 1;
                int sb_kor = headers.IndexOf("SB korisnik") + 1;
                int prof_napl_sb = headers.IndexOf("Profil naplate SB korisnika") + 1;
                int ser_sim = headers.IndexOf("Serijski broj SIM-a") + 1;
                int dat_poc_uo = headers.IndexOf("Datum početka ugovorne obveze") + 1;
                int dat_kraj_uo = headers.IndexOf("Datum isteka ugovorne obveze") + 1;
                int pnp = headers.IndexOf("Skraćeni broj (PNP)") + 1;
                int odl_prof = headers.IndexOf("VPN odlazni profil") + 1;
                int dol_prof = headers.IndexOf("VPN dolazni profil") + 1;
                int vpn_budget = headers.IndexOf("Iznos limita - VPN Budget") + 1;
                int limit = headers.IndexOf("Iznos limita potrošnje") + 1;
                int korp_apn = headers.IndexOf("Korporativni APN") + 1;
                int multisim = headers.IndexOf("MultiSIM nominacija") + 1;
                int vrsta_usluge = headers.IndexOf("Vrsta usluge") + 1;
                int vrsta_proiz = headers.IndexOf("Vrsta proizvoda") + 1;
                int klas_proiz = headers.IndexOf("Klasifikacija proizvoda") + 1;
                int stat_uo = headers.IndexOf("Status ugovorne obveze") + 1;
                int br_dana_uo = headers.IndexOf("Preostalo dana ugovorne obveze") + 1;

                int multisimcount = 0;
                int korporativniAPN = 0;
                int limitPotrosnje = 0;

                string temp_broj = _2dArray[2, broj_telefona].ToString();

                for (var i = 2; i <= rowCount; i++)
                {
                    if (String.Compare(_2dArray[i, status].ToString(), "Active") == 0 || String.Compare(_2dArray[i, status].ToString(), "Suspended") == 0)
                    {
                        if (String.Compare(temp_broj, _2dArray[i, broj_telefona].ToString()) == 0)
                        {
                            Proizvodi row_proizvod = new Proizvodi();
                            Proizvodi unique_proizvod = new Proizvodi();

                            if (_2dArray[i, kor_naplata] != null && String.Compare(_2dArray[i, kor_naplata].ToString(), "") != 0 && String.Compare(_2dArray[i, kor_naplata].ToString(), "--") != 0)
                                row_siebel_export.KorisnikZaNaplatu = _2dArray[i, kor_naplata].ToString();
                            if (_2dArray[i, kor_usluga] != null && String.Compare(_2dArray[i, kor_usluga].ToString(), "") != 0 && String.Compare(_2dArray[i, kor_usluga].ToString(), "--") != 0)
                                row_siebel_export.KorisnikZaUslugu = _2dArray[i, kor_usluga].ToString();
                            if (_2dArray[i, broj_telefona] != null && String.Compare(_2dArray[i, broj_telefona].ToString(), "") != 0 && String.Compare(_2dArray[i, broj_telefona].ToString(), "--") != 0)
                                row_siebel_export.BrojTelefona = _2dArray[i, broj_telefona].ToString();
                            if (_2dArray[i, status] != null && String.Compare(_2dArray[i, status].ToString(), "") != 0 && String.Compare(_2dArray[i, status].ToString(), "--") != 0)
                                row_siebel_export.Status = _2dArray[i, status].ToString();
                            if (_2dArray[i, prof_napl] != null && String.Compare(_2dArray[i, prof_napl].ToString(), "") != 0 && String.Compare(_2dArray[i, prof_napl].ToString(), "--") != 0)
                                row_siebel_export.ProfilNaplate = _2dArray[i, prof_napl].ToString();
                            if (_2dArray[i, dat_poc_uo] != null && String.Compare(_2dArray[i, dat_poc_uo].ToString(), "") != 0 && String.Compare(_2dArray[i, dat_poc_uo].ToString(), "--") != 0)
                            {
                                try
                                {
                                    if (Convert.ToDateTime(row_siebel_export.PocetakUO) < Convert.ToDateTime(DateTime.FromOADate(Convert.ToDouble(_2dArray[i, dat_poc_uo].ToString())).ToString("dd.MM.yyyy")))
                                        row_siebel_export.PocetakUO = DateTime.FromOADate(Convert.ToDouble(_2dArray[i, dat_poc_uo].ToString())).ToString("dd.MM.yyyy");
                                }
                                catch
                                {
                                    if (Convert.ToDateTime(row_siebel_export.PocetakUO) < Convert.ToDateTime(_2dArray[i, dat_poc_uo].ToString().Substring(2, _2dArray[i, dat_poc_uo].ToString().Length - 2)))
                                        row_siebel_export.PocetakUO = DateTime.ParseExact(_2dArray[i, dat_poc_uo].ToString().Substring(2, _2dArray[i, dat_poc_uo].ToString().Length - 2), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd.MM.yyyy");
                                }
                            }
                            if (_2dArray[i, dat_kraj_uo] != null && String.Compare(_2dArray[i, dat_kraj_uo].ToString(), "") != 0 && String.Compare(_2dArray[i, dat_kraj_uo].ToString(), "--") != 0)
                            {
                                try
                                {
                                    if (Convert.ToDateTime(row_siebel_export.IstekUO) < Convert.ToDateTime(DateTime.FromOADate(Convert.ToDouble(_2dArray[i, dat_kraj_uo].ToString())).ToString("dd.MM.yyyy")))
                                        row_siebel_export.IstekUO = DateTime.FromOADate(Convert.ToDouble(_2dArray[i, dat_kraj_uo].ToString())).ToString("dd.MM.yyyy");
                                }
                                catch
                                {
                                    if (Convert.ToDateTime(row_siebel_export.IstekUO) < Convert.ToDateTime(_2dArray[i, dat_kraj_uo].ToString().Substring(2, _2dArray[i, dat_kraj_uo].ToString().Length - 2)))
                                        row_siebel_export.IstekUO = DateTime.ParseExact(_2dArray[i, dat_kraj_uo].ToString().Substring(2, _2dArray[i, dat_kraj_uo].ToString().Length - 2), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd.MM.yyyy");
                                }
                            }
                            if (_2dArray[i, pnp] != null && String.Compare(_2dArray[i, pnp].ToString(), "") != 0 && String.Compare(_2dArray[i, pnp].ToString(), "--") != 0)
                                row_siebel_export.PNP = _2dArray[i, pnp].ToString();
                            if (_2dArray[i, odl_prof] != null && String.Compare(_2dArray[i, odl_prof].ToString(), "") != 0 && String.Compare(_2dArray[i, odl_prof].ToString(), "--") != 0)
                                row_siebel_export.OdlazniProfil = _2dArray[i, odl_prof].ToString();
                            if (_2dArray[i, dol_prof] != null && String.Compare(_2dArray[i, dol_prof].ToString(), "") != 0 && String.Compare(_2dArray[i, dol_prof].ToString(), "--") != 0)
                                row_siebel_export.DolazniProfil = _2dArray[i, dol_prof].ToString();
                            if (_2dArray[i, stat_uo] != null && String.Compare(_2dArray[i, stat_uo].ToString(), "") != 0 && String.Compare(_2dArray[i, stat_uo].ToString(), "--") != 0)
                                row_siebel_export.StatusUgovorneObveze = _2dArray[i, stat_uo].ToString();
                            if (_2dArray[i, br_dana_uo] != null && String.Compare(_2dArray[i, br_dana_uo].ToString(), "") != 0 && String.Compare(_2dArray[i, br_dana_uo].ToString(), "--") != 0)
                                row_siebel_export.PreostaloDana = _2dArray[i, br_dana_uo].ToString();

                            
                            if (_2dArray[i, multisim] != null && String.Compare(_2dArray[i, multisim].ToString(), "") != 0 && String.Compare(_2dArray[i, multisim].ToString(), "--") != 0)
                            {
                                row_siebel_export.MultiSIM_nominacija = _2dArray[i, multisim].ToString();
                                multisimcount = 1;
                            }
                            if (_2dArray[i, korp_apn] != null && String.Compare(_2dArray[i, korp_apn].ToString(), "") != 0 && String.Compare(_2dArray[i, korp_apn].ToString(), "--") != 0)
                            {
                                row_siebel_export.KorporativniAPN = _2dArray[i, korp_apn].ToString();
                                korporativniAPN = 1;
                            }
                            if (_2dArray[i, limit] != null && String.Compare(_2dArray[i, limit].ToString(), "") != 0 && String.Compare(_2dArray[i, limit].ToString(), "--") != 0)
                            {
                                row_siebel_export.LimitPotrosnje = _2dArray[i, limit].ToString();
                                limitPotrosnje = 1;
                            }

                            if (_2dArray[i, klas_proiz] == null)
                            {
                                continue;
                            }
                            else if (String.Compare(_2dArray[i, klas_proiz].ToString(), "Root Service") == 0)
                            {
                                if (_2dArray[i, dat_akt] != null && String.Compare(_2dArray[i, dat_akt].ToString(), "") != 0 && String.Compare(_2dArray[i, dat_akt].ToString(), "--") != 0)
                                {
                                    try
                                    {
                                        row_siebel_export.DatumAktivacijeUsluge = DateTime.FromOADate(Convert.ToDouble(_2dArray[i, dat_akt].ToString())).ToString("dd.MM.yyyy");
                                    }
                                    catch
                                    {
                                        row_siebel_export.DatumAktivacijeUsluge = DateTime.ParseExact(_2dArray[i, dat_akt].ToString().Substring(2, _2dArray[i, dat_akt].ToString().Length - 2), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd.MM.yyyy");
                                    }
                                }
                                if (_2dArray[i, proizv] != null && String.Compare(_2dArray[i, proizv].ToString(), "") != 0 && String.Compare(_2dArray[i, proizv].ToString(), "--") != 0)
                                    row_siebel_export.Usluga = _2dArray[i, proizv].ToString();
                            }
                            else if (String.Compare(_2dArray[i, klas_proiz].ToString(), "Tariff") == 0)
                            {
                                if (_2dArray[i, dat_akt] != null && String.Compare(_2dArray[i, dat_akt].ToString(), "") != 0 && String.Compare(_2dArray[i, dat_akt].ToString(), "--") != 0)
                                {
                                    try
                                    {
                                        row_siebel_export.DatumAktivacijeTarife = DateTime.FromOADate(Convert.ToDouble(_2dArray[i, dat_akt].ToString())).ToString("dd.MM.yyyy");
                                    }
                                    catch
                                    {
                                        row_siebel_export.DatumAktivacijeTarife = DateTime.ParseExact(_2dArray[i, dat_akt].ToString().Substring(2, _2dArray[i, dat_akt].ToString().Length - 2), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd.MM.yyyy");
                                    }
                                }
                                if (_2dArray[i, proizv] != null && String.Compare(_2dArray[i, proizv].ToString(), "") != 0 && String.Compare(_2dArray[i, proizv].ToString(), "--") != 0)
                                    row_siebel_export.Tarifa = _2dArray[i, proizv].ToString();
                            }
                            else if (String.Compare(_2dArray[i, klas_proiz].ToString(), "Hardware") == 0)
                            {
                                /* simovi lista */
                                SIM_S sim = new SIM_S();
                                if (_2dArray[i, proizv] != null && String.Compare(_2dArray[i, proizv].ToString(), "") != 0 && String.Compare(_2dArray[i, proizv].ToString(), "--") != 0)
                                    sim.Naziv = _2dArray[i, proizv].ToString();
                                if (_2dArray[i, ser_sim] != null && String.Compare(_2dArray[i, ser_sim].ToString(), "") != 0 && String.Compare(_2dArray[i, ser_sim].ToString(), "--") != 0)
                                    sim.Serial = _2dArray[i, ser_sim].ToString();
                                row_siebel_export.Simovi.Add(sim);
                                if (unique_sims.Find(x => String.Compare(x, sim.Naziv) == 0) == null)
                                {
                                    unique_sims.Add(sim.Naziv);
                                }
                            }
                            else if (String.Compare(_2dArray[i, klas_proiz].ToString(), "Split Biller") == 0)
                            {
                                if (_2dArray[i, sb_kor] != null && String.Compare(_2dArray[i, sb_kor].ToString(), "") != 0 && String.Compare(_2dArray[i, sb_kor].ToString(), "--") != 0)
                                    row_siebel_export.SplitBiller = _2dArray[i, sb_kor].ToString();
                                if (_2dArray[i, vpn_budget] != null && String.Compare(_2dArray[i, vpn_budget].ToString(), "") != 0 && String.Compare(_2dArray[i, vpn_budget].ToString(), "--") != 0)
                                    row_siebel_export.Vpn_budget = _2dArray[i, vpn_budget].ToString();
                                if (_2dArray[i, prof_napl_sb] != null && String.Compare(_2dArray[i, prof_napl_sb].ToString(), "") != 0 && String.Compare(_2dArray[i, prof_napl_sb].ToString(), "--") != 0)
                                    row_siebel_export.ProfilNaplateSB = _2dArray[i, prof_napl_sb].ToString();
                                splitBillerActive = true;
                            }
                            else
                            {
                                //SVE OSTALO
                                if (_2dArray[i, dat_akt] != null && String.Compare(_2dArray[i, dat_akt].ToString(), "") != 0 && String.Compare(_2dArray[i, dat_akt].ToString(), "--") != 0)
                                {
                                    try
                                    {
                                        if (Convert.ToDateTime(row_proizvod.DatumAktivacije) < Convert.ToDateTime(DateTime.FromOADate(Convert.ToDouble(_2dArray[i, dat_akt].ToString())).ToString("dd.MM.yyyy")))
                                            row_proizvod.DatumAktivacije = DateTime.FromOADate(Convert.ToDouble(_2dArray[i, dat_akt].ToString())).ToString("dd.MM.yyyy");
                                    }
                                    catch
                                    {
                                        if (Convert.ToDateTime(row_proizvod.DatumAktivacije) < Convert.ToDateTime(_2dArray[i, dat_akt].ToString().Substring(2, _2dArray[i, dat_akt].ToString().Length - 2)))
                                            row_proizvod.DatumAktivacije = DateTime.ParseExact(_2dArray[i, dat_akt].ToString().Substring(2, _2dArray[i, dat_akt].ToString().Length - 2), "dd-MM-yyyy", CultureInfo.InvariantCulture).ToString("dd.MM.yyyy");
                                    
                                    }
                                }
                                if (_2dArray[i, proizv] != null && String.Compare(_2dArray[i, proizv].ToString(), "") != 0 && String.Compare(_2dArray[i, proizv].ToString(), "--") != 0)
                                {
                                    row_proizvod.Proizvod = _2dArray[i, proizv].ToString();
                                    unique_proizvod.Proizvod = _2dArray[i, proizv].ToString();
                                }
                                if (_2dArray[i, klas_proiz] != null && String.Compare(_2dArray[i, klas_proiz].ToString(), "") != 0 && String.Compare(_2dArray[i, klas_proiz].ToString(), "--") != 0)
                                {
                                    row_proizvod.KlasifikacijaProizvoda = _2dArray[i, klas_proiz].ToString();
                                    unique_proizvod.KlasifikacijaProizvoda = _2dArray[i, klas_proiz].ToString();
                                }
                                row_siebel_export.Proizvodi.Add(row_proizvod);

                                if (unique_proizvodi.Find(x => String.Compare(x.Proizvod, unique_proizvod.Proizvod) == 0) == null)
                                {
                                    unique_proizvodi.Add(unique_proizvod);
                                }
                            }

                            if (i != rowCount && String.Compare(temp_broj, _2dArray[i + 1, broj_telefona].ToString()) != 0)
                            {
                                row_siebel_export.Simovi = row_siebel_export.Simovi.OrderBy(x => x.Naziv).ToList();
                                dataSiebelExport.Add(row_siebel_export);
                                
                                if (i + 1 < rowCount)
                                {
                                    temp_broj = _2dArray[i + 1, broj_telefona].ToString();
                                    row_siebel_export = new Siebel_export();
                                }

                            }
                        }
                    }
                    else if (i+1 <= rowCount &&  String.Compare(temp_broj, _2dArray[i + 1, broj_telefona].ToString()) != 0)
                    {
                        if (row_siebel_export.Status != null)
                        {
                            row_siebel_export.Simovi = row_siebel_export.Simovi.OrderBy(x => x.Naziv).ToList();
                            dataSiebelExport.Add(row_siebel_export);
                        }
                        if (i + 1 < rowCount)
                        {
                            temp_broj = _2dArray[i + 1, broj_telefona].ToString();
                            row_siebel_export = new Siebel_export();
                        }

                    }

                    if (i == rowCount && row_siebel_export.BrojTelefona != null)
                    {
                        row_siebel_export.Simovi = row_siebel_export.Simovi.OrderBy(x => x.Naziv).ToList();
                        dataSiebelExport.Add(row_siebel_export);
                    }
                }

                unique_proizvodi = unique_proizvodi.OrderBy(x => x.KlasifikacijaProizvoda).ThenBy(x => x.Proizvod).ToList();
              
                unique_sims = unique_sims.OrderBy(x => x).ToList();
                
                dataSiebelExport = dataSiebelExport.OrderBy(x => x.BrojTelefona).ToList();

                xlWorkbook.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();

                loadInfo.Text = "Podatci su učitani";

                int rowCountExport = dataSiebelExport.Count + 1;



                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
                Excel._Worksheet workSheet = workBook.Worksheets[1];
                workSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;


                object[,] _2dData = new object[rowCountExport, 16 + multisimcount + korporativniAPN + limitPotrosnje + 3 + unique_sims.Count + unique_proizvodi.Count];

                workSheet.Name = "Export";


                
               

                _2dData[0, 0] = "Korisnik za naplatu";
                _2dData[0, 1] = "Korisnik za uslugu";
                _2dData[0, 2] = "Datum aktivacije usluge";
                _2dData[0, 3] = "Usluga";
                _2dData[0, 4] = "Broj telefona";
                _2dData[0, 5] = "Status";
                _2dData[0, 6] = "Profil naplate";
                if (splitBillerActive)
                {
                    _2dData[0, 7] = "Split Biller";
                    _2dData[0, 8] = "Iznos limita - VPN Budget";
                    _2dData[0, 9] = "Profil Naplate SB";
                    sbBroj = 3;
                }
                _2dData[0, 7 + sbBroj] = "Datum aktivacije tarife";
                _2dData[0, 8 + sbBroj] = "Tarifa";
                _2dData[0, 9 + sbBroj] = "Početak ugovorne obveze";
                _2dData[0, 10 + sbBroj] = "Istek ugovorne obveze";
                _2dData[0, 11 + sbBroj] = "PNP";
                _2dData[0, 12 + sbBroj] = "Odlazni profil";
                _2dData[0, 13 + sbBroj] = "Dolazni profil";
                _2dData[0, 14 + sbBroj] = "Status ugovorne obveze";
                _2dData[0, 15 + sbBroj] = "Preostali broj dana ugovorne obveze";

                if (limitPotrosnje == 1)
                {
                    _2dData[0, 16 + sbBroj] = "Iznos limita potrošnje";
                }
                if (korporativniAPN == 1)
                {
                    _2dData[0, 16 + limitPotrosnje + sbBroj] = "Korporativni APN";
                }
                if (multisimcount == 1)
                {
                    _2dData[0, 16 + limitPotrosnje + korporativniAPN + sbBroj] = "MultiSIM nominacija";
                }

                string stupac = GetExcelColumnName(16 + multisimcount + limitPotrosnje + korporativniAPN + sbBroj + unique_proizvodi.Count + unique_sims.Count);
                workSheet.get_Range("a1", stupac + "1").Cells.Interior.Color = System.Drawing.Color.Orange;
                workSheet.get_Range("a1", stupac + "1").Cells.Font.Color = System.Drawing.Color.Black;
                workSheet.get_Range("a1", stupac + "1").Cells.Font.Bold = true;
                workSheet.get_Range("a1", stupac + "1").Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3d;
                
                Excel.Range c1 = workSheet.Cells[1, 17 + multisimcount + limitPotrosnje + korporativniAPN + sbBroj];
                Excel.Range c2 = workSheet.Cells[rowCountExport, 16 + multisimcount + limitPotrosnje + korporativniAPN + sbBroj + unique_sims.Count];
                Excel.Range rangeNum = workSheet.get_Range(c1, c2);
                rangeNum.NumberFormat = "@";
                rangeNum.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                
                for (int k = 1; k <= unique_sims.Count; k++)
                {
                    _2dData[0, 15 + multisimcount + limitPotrosnje + korporativniAPN + sbBroj + k] = unique_sims[k - 1];
                }
                for (int k = 1; k <= unique_proizvodi.Count; k++)
                {
                    _2dData[0, 15 + multisimcount + limitPotrosnje + korporativniAPN + sbBroj + unique_sims.Count + k] = unique_proizvodi[k - 1].Proizvod;
                }
        


                for (int i = 0; i < dataSiebelExport.Count; i++)
                {
                    _2dData[i + 1, 0] = dataSiebelExport[i].KorisnikZaNaplatu;
                    _2dData[i + 1, 1] = dataSiebelExport[i].KorisnikZaUslugu;
                    _2dData[i + 1, 2] = Convert.ToDateTime(dataSiebelExport[i].DatumAktivacijeUsluge);
                    _2dData[i + 1, 3] = dataSiebelExport[i].Usluga;
                    _2dData[i + 1, 4] = dataSiebelExport[i].BrojTelefona;
                    _2dData[i + 1, 5] = dataSiebelExport[i].Status;
                    _2dData[i + 1, 6] = dataSiebelExport[i].ProfilNaplate;
                    if (splitBillerActive)
                    {
                        _2dData[i + 1, 7] = dataSiebelExport[i].SplitBiller;
                        _2dData[i + 1, 8] = dataSiebelExport[i].Vpn_budget;
                        _2dData[i + 1, 9] = dataSiebelExport[i].ProfilNaplateSB;
                    }
                    _2dData[i + 1, 7 + sbBroj] = Convert.ToDateTime(dataSiebelExport[i].DatumAktivacijeTarife);
                    _2dData[i + 1, 8 + sbBroj] = dataSiebelExport[i].Tarifa;
                    _2dData[i + 1, 9 + sbBroj] = Convert.ToDateTime(dataSiebelExport[i].PocetakUO);
                    _2dData[i + 1, 10 + sbBroj] = Convert.ToDateTime(dataSiebelExport[i].IstekUO);
                    _2dData[i + 1, 11 + sbBroj] = dataSiebelExport[i].PNP;
                    _2dData[i + 1, 12 + sbBroj] = dataSiebelExport[i].OdlazniProfil;
                    _2dData[i + 1, 13 + sbBroj] = dataSiebelExport[i].DolazniProfil;
                    _2dData[i + 1, 14 + sbBroj] = dataSiebelExport[i].StatusUgovorneObveze;
                    _2dData[i + 1, 15 + sbBroj] = dataSiebelExport[i].PreostaloDana;

                    if (limitPotrosnje == 1)
                    {
                        _2dData[i + 1, 16 + sbBroj] = dataSiebelExport[i].LimitPotrosnje;
                    }
                    if (korporativniAPN == 1)
                    {
                        _2dData[i + 1, 16 + limitPotrosnje + sbBroj] = dataSiebelExport[i].KorporativniAPN;
                    }                    
                    if (multisimcount == 1)
                    {
                        _2dData[i + 1, 16 + limitPotrosnje + korporativniAPN + sbBroj] = dataSiebelExport[i].MultiSIM_nominacija;
                    }
                    foreach (SIM_S temp in dataSiebelExport[i].Simovi)
                    {
                        _2dData[i + 1, 16 + multisimcount + limitPotrosnje + korporativniAPN + sbBroj + unique_sims.IndexOf(temp.Naziv)] = temp.Serial;
                    }

                    foreach (Proizvodi temp in dataSiebelExport[i].Proizvodi)
                    {
                        int index = unique_proizvodi.FindIndex(x => String.Compare(x.Proizvod, temp.Proizvod) == 0);
                        _2dData[i + 1, 16 + multisimcount + limitPotrosnje + korporativniAPN + sbBroj + unique_sims.Count + index] = "X";
                    }
                }

                c1 = workSheet.Cells[1, 1];
                c2 = workSheet.Cells[rowCountExport, 16 + multisimcount + limitPotrosnje + korporativniAPN + sbBroj + unique_sims.Count + unique_proizvodi.Count];
                Excel.Range range = workSheet.get_Range(c1, c2);

                if (chkDatumi.Checked)
                {
                    for (int i = 0; i < dataSiebelExport.Count; i++)
                    {
                        foreach (Proizvodi temp in dataSiebelExport[i].Proizvodi)
                        {
                            int index = unique_proizvodi.FindIndex(x => String.Compare(x.Proizvod, temp.Proizvod) == 0);
                            c1 = workSheet.Cells[i + 2, 17 + multisimcount + limitPotrosnje + korporativniAPN + sbBroj + unique_sims.Count + index];
                            c1.AddComment(temp.DatumAktivacije);
                        }
                    }
                }
                
                c1 = workSheet.Cells[1, 5];
                c2 = workSheet.Cells[rowCountExport, 7];
                rangeNum = workSheet.get_Range(c1, c2);
                rangeNum.NumberFormat = "#";
                rangeNum.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                c1 = workSheet.Cells[1, 12];
                c2 = workSheet.Cells[rowCountExport, 12];
                workSheet.get_Range(c1, c2).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                c1 = workSheet.Cells[1, 16];
                c2 = workSheet.Cells[rowCountExport, 16 + sbBroj + unique_proizvodi.Count + unique_sims.Count];
                workSheet.get_Range(c1, c2).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
               
                range.Value = _2dData;
                workSheet.Application.ActiveWindow.SplitRow = 1;
                workSheet.Application.ActiveWindow.FreezePanes = true;
                Excel.Range firstRow = (Excel.Range)workSheet.Rows[1];
                firstRow.Activate();
                firstRow.Select();
                firstRow.AutoFilter(1,
                                    Type.Missing,
                                    Excel.XlAutoFilterOperator.xlAnd,
                                    Type.Missing,
                                    true);
                workSheet.get_Range("A:" + stupac, Type.Missing).Columns.AutoFit();
                excelApp.DisplayAlerts = true;
                excelPath = Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
                bool tempSave = false;
                int saveCounter = 0;
                while(!tempSave)
                {
                    try
                    {
                        tempSave = true;
                        if (saveCounter == 0)
                        {
                            workSheet.SaveAs(excelPath + "\\" + "export " + dataSiebelExport[0].KorisnikZaNaplatu + " " + DateTime.Now.ToShortDateString() + ".xlsx");
                        }
                        else
                        {
                            workSheet.SaveAs(excelPath + "\\" + "export " + dataSiebelExport[0].KorisnikZaNaplatu + " " + DateTime.Now.ToShortDateString() + ".xlsx");
                        }
                       
                    }
                    catch 
                    {
                        tempSave = false;
                        saveCounter++;
                    }
                }
                workBook.Close(true, Type.Missing, Type.Missing);
                excelApp.Quit();
            

                Cursor = DefaultCursor;
                loadInfo.Text = "Export završen";
                loadInfo.ForeColor = Color.Green;
                chkDatumi.Enabled = false;
                btnExport.Text = "Zatvori";
                btnOpenExport.Enabled = false;
            }
        }

        public string excelPath = "";
        private void openExport_Click(object sender, EventArgs e)
        {
            if (String.Compare(btnOpenExport.Text, "Isprazni") == 0) {
                openExportDialog.Dispose();
                loadInfo.Text = "Nema podataka";
                loadInfo.ForeColor = Color.Red;
                btnExport.Enabled = false;
                btnOpenExport.Text = "Učitaj export";
            }
            else if (String.Compare(openExportDialog.ShowDialog().ToString(), "Cancel") == 0)
            {
                openExportDialog.Dispose();
                btnExport.Enabled = false;
            }
            else
            {
                FileInfo temp = new FileInfo(openExportDialog.FileName);
                if (File.Exists(temp.FullName))
                {
                    excelPath = temp.FullName;
                    chkDatumi.Enabled = true;
                    loadInfo.Text = "Pronađeni podatci";
                    btnOpenExport.Text = "Isprazni";
                    loadInfo.ForeColor = Color.Black;
                    btnExport.Enabled = true;
                }
                else
                {
                    loadInfo.Text = "Nema podataka";
                    loadInfo.ForeColor = Color.Red;
                    btnExport.Enabled = false;
                }
            }
        }

        private void Export_Load(object sender, EventArgs e)
        {
            loadInfo.Font = new Font("Futura Bk CE", 14, FontStyle.Regular);
            btnOpenExport.Font = new Font("Futura Bk CE", 14, FontStyle.Regular);
            btnExport.Font = new Font("Futura Bk CE", 18, FontStyle.Regular);
            chkDatumi.Font = new Font("Futura Bk CE", 11, FontStyle.Regular);
            lblVersion.Text = "v " + ProductVersion;
        }

        private void chkDatumi_CheckedChanged(object sender, EventArgs e)
        {
            if (chkDatumi.Checked  == true && MessageBox.Show("Ispisom datuma u komentarima se produljuje vrijeme generiranja izvoza tablice.", "Upozorenje", MessageBoxButtons.YesNo) == DialogResult.No)
            {
                chkDatumi.Checked = false;
            }
        }
    }
}
