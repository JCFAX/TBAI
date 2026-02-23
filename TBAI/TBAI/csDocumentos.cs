using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TBAI
{
    public class csDocumentos
    {
        public int DocEntry;
        public string ObjType;
        public string DocNum;
        public string CardCode;
        public string CardName;
        public DateTime DocDate;
        public string NumAtCard;
        public string DocType;
        public string Country;
        public string Province;
        public string City;
        public string Address;
        public string ZipCode;
        public string U_B1SYS_INV_TYPE;
        public double DocTotal;
        public string DocCur;
        public string LicTradNum;
        public string CodigoEnvio;
        public string DocEntryR;
        public string InvoiceType;
        public string Enterprise;
        public string AX_SAPSerieName;
        public string AX_DocNum;
        public string AX_ClaveRegimenIvaOpTrascendencia;
        public string AX_NombreOperacion;
        public string AX_TBAI_FECasSuj;
        public string AX_TBAI_FECasExe;
        public string AX_TBAI_FECauExe;
        public string AX_TBAI_FETipNExe;
        public string AX_TBAI_FECauNSuj;
        public string AX_TBAI_FEClaRegEsp;
        public string AX_TBAI_FRClaRegEsp;
        public string AX_TBAI_Ser;
        public string AX_TBAI_InvSujPas;
        public string AX_TBAI_BieInv;
        public string AX_TBAI_DetOpeInt;
        public string AX_TBAI_Imp;
        public string AX_TBAI_FacSim;
        public string AX_TBAI_FacRegSim;
        public string AX_TBAI_FacRegRecEqu;
        public string AX_TBAI_CasREAV;
        public string AX_TBAI_FRCG240;
        public string AX_TBAI_Arr;
        public string AX_TBAI_CasREGEAju;

        public List<csDocumentos_Lineas> DocumentLines;
    }

    public class csDocumentos_Lineas
    {
        public int DocEntry;
        public int LineNum;
        public string ItemCode;
        public string Dscription;
        public double Quantity;
        public double VatSum;
        public double LineTotal;
        public string ResidenNum;
        public string AbsEntry;
        public string LineSeq;
        public double VatPercent;
        public double NdPercent;
        public double EqPercent;
        public double baseSum;
        public double DeductSum;
        public double EqSum;
        public string CrditDebit;
        public string TaxCode;
        public string Description;
        public string U_AX_TBAI_CodImp;
        public string NombreOperacion;
        public string U_AX_TBAI_FECasSuj;
        public string U_AX_TBAI_FECasExe;
        public string U_AX_TBAI_FECauExe;
        public string U_AX_TBAI_FETipNExe;
        public string U_AX_TBAI_FECauNSuj;
        public string U_AX_TBAI_FEClaRegEsp;
        public string U_AX_TBAI_FRClaRegEsp;
        public string U_AX_TBAI_Ser;
        public string U_AX_TBAI_InvSujPas;
        public string U_AX_TBAI_BieInv;
        public string U_AX_TBAI_DetOpeInt;
        public string U_AX_TBAI_Imp;
        public string U_AX_TBAI_FacSim;
        public string U_AX_TBAI_FacRegSim;
        public string U_AX_TBAI_FacRegRecEqu;
        public string U_AX_TBAI_CasREAV;
        public string U_AX_TBAI_FRCG240;
        public string U_AX_TBAI_Arr;
        public string U_AX_TBAI_CasREGEAju;
    }

    public class csListaDocumentos
    {
        public List<csDocumentos> ListaDocumentos { get; set; }
    }

    public class csDocumentosQR
    {
        public int DocEntry;
        public string ObjType;
        public string DocNum;
        public string Serie;
        public string Code;
        public string QR;
        public byte[] QR_64;
        public string DocumentoTBAIFirmado;
    }
}
