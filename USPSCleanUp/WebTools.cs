using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace USPSCleanUp
{
    public class WebTools
    {
        //private const string BaseURL = "http://testing.shippingapis.com/ShippingAPITest.dll";
        //private const string BaseURL = "http://production.shippingapis.com/ShippingAPITest.dll?API=Verify";
        private const string BaseURL = "http://production.shippingapis.com/ShippingAPITest.dll";
        //Web client instance.

        private WebClient wsClient = new WebClient();

        //User ID obtained from USPS.

        public string USPS_UserID = "607FRIEN1074";

        public WebTools()

        {

        }



        //Constructor with User ID parameter.

        public WebTools(string New_UserID)

        {

            USPS_UserID = New_UserID;

        }

        private string GetDataFromSite(string USPS_Request)

        {

            string strResponse = "";



            //Send the request to USPS.

            byte[] ResponseData = wsClient.DownloadData(USPS_Request);

            //Convert byte stream to string data.

            foreach (byte oItem in ResponseData)

                strResponse += (char)oItem;



            return strResponse;

        }
        public string AddressValidateRequest(string Address1,

                                     string Address2,

                                     string City,

                                     string State,

                                     string Zip5,

                                     string Zip4)

        {

            //http://production.shippingapis.com/ShippingAPITest.dll?API=Verify

            //  &XML=<AddressValidateRequest USERID="xxxxxxx"><Address ID="0"><Address1></Address1>

            //  <Address2>6406 Ivy Lane</Address2><City>Greenbelt</City><State>MD</State>

            //  <Zip5></Zip5><Zip4></Zip4></Address></AddressValidateRequest>



            string strResponse = "", strUSPS = "";



            strUSPS = BaseURL + "?API=Verify&XML=<AddressValidateRequest USERID=\"" + USPS_UserID + "\">";

            strUSPS += "<Address ID=\"0\">";

            strUSPS += "<Address1>" + Address1 + "</Address1>";

            strUSPS += "<Address2>" + Address2 + "</Address2>";

            strUSPS += "<City>" + City + "</City>";

            strUSPS += "<State>" + State + "</State>";

            strUSPS += "<Zip5>" + Zip5 + "</Zip5>";

            strUSPS += "<Zip4>" + Zip4 + "</Zip4>";

            strUSPS += "</Address></AddressValidateRequest>";



            //Send the request to USPS.

            strResponse = GetDataFromSite(strUSPS);



            return strResponse;

        }
    }
}
