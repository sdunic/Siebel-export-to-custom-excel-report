using System;
using System.Collections.Generic;
using System.Linq;

namespace Siebel_custom_report
{

    /*
        Barring
        Hardware
        Main Service
        Option Group
        Options
        VAS
     */

    class Siebel_export
    {
        private string korisnikZaNaplatu;
        private string korisnikZaUslugu;
        private string brojTelefona;
        private string status;
        private string profilNaplate;
        private string pocetakUO;
        private string istekUO;
        private string pNP;
        private string odlazniProfil;
        private string dolazniProfil;
        private string statusUgovorneObveze;
        private string preostaloDana;
        private string usluga;
        private string datumAktivacijeUsluge;
        private string tarifa;
        private string datumAktivacijeTarife;
        private string splitBiller;
        private string profilNaplateSB;
        private string vpn_budget;
        private string multiSIM_nominacija;
        private string limitPotrosnje;
        private string korporativniAPN;

        public string KorporativniAPN
        {
            get { return korporativniAPN; }
            set { korporativniAPN = value; }
        }


        public string LimitPotrosnje
        {
            get { return limitPotrosnje; }
            set { limitPotrosnje = value; }
        }


        public string MultiSIM_nominacija
        {
            get { return multiSIM_nominacija; }
            set { multiSIM_nominacija = value; }
        }

        public string Vpn_budget
        {
            get { return vpn_budget; }
            set { vpn_budget = value; }
        }

        public string ProfilNaplateSB
        {
            get { return profilNaplateSB; }
            set { profilNaplateSB = value; }
        }

        public string SplitBiller
        {
            get { return splitBiller; }
            set { splitBiller = value; }
        }

        public string DatumAktivacijeTarife
        {
            get { return datumAktivacijeTarife; }
            set { datumAktivacijeTarife = value; }
        }

        public List<Proizvodi> Proizvodi = new List<Proizvodi>();
        public List<SIM_S> Simovi = new List<SIM_S>();

        public string Tarifa
        {
            get { return tarifa; }
            set { tarifa = value; }
        }

        public string DatumAktivacijeUsluge
        {
            get { return datumAktivacijeUsluge; }
            set { datumAktivacijeUsluge = value; }
        }

        public string Usluga
        {
            get { return usluga; }
            set { usluga = value; }
        }
        
        public string PreostaloDana
        {
            get { return preostaloDana; }
            set { preostaloDana = value; }
        }

        public string StatusUgovorneObveze
        {
            get { return statusUgovorneObveze; }
            set { statusUgovorneObveze = value; }
        }

        public string DolazniProfil
        {
            get { return dolazniProfil; }
            set { dolazniProfil = value; }
        }

        public string OdlazniProfil
        {
            get { return odlazniProfil; }
            set { odlazniProfil = value; }
        }

        public string PNP
        {
            get { return pNP; }
            set { pNP = value; }
        }

        public string IstekUO
        {
            get { return istekUO; }
            set { istekUO = value; }
        }

        public string PocetakUO
        {
            get { return pocetakUO; }
            set { pocetakUO = value; }
        }

        public string ProfilNaplate
        {
            get { return profilNaplate; }
            set { profilNaplate = value; }
        }

        public string Status
        {
            get { return status; }
            set { status = value; }
        }

        public string BrojTelefona
        {
            get { return brojTelefona; }
            set { brojTelefona = value; }
        }

        public string KorisnikZaUslugu
        {
            get { return korisnikZaUslugu; }
            set { korisnikZaUslugu = value; }
        }

        public string KorisnikZaNaplatu
        {
            get { return korisnikZaNaplatu; }
            set { korisnikZaNaplatu = value; }
        }
    }

    class Proizvodi
    {
        private string klasifikacijaProizvoda;

        public string KlasifikacijaProizvoda
        {
            get { return klasifikacijaProizvoda; }
            set { klasifikacijaProizvoda = value; }
        }
        private string proizvod;

        public string Proizvod
        {
            get { return proizvod; }
            set { proizvod = value; }
        }
        private string datumAktivacije;

        public string DatumAktivacije
        {
            get { return datumAktivacije; }
            set { datumAktivacije = value; }
        }
    }

    class SIM_S
    {
        private string naziv;
        private string serial;
 
        public string Serial
        {
            get { return serial; }
            set { serial = value; }
        }

        public string Naziv
        {
            get { return naziv; }
            set { naziv = value; }
        }
    }
}
