using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;

namespace Sistema_Indicadores.Clases
{
    public class Notificaciones
    {
        public string SendNotificationJSON(string title, string body)
        {
            try
            {
            string SERVER_API_KEY = "AIzaSyANuakkocCPEoISIJG7DID6StVrEu3Dm80";
            string SENDER_ID = "47398042328";
            string deviceId = "bda93f56ce507588";

                WebRequest tRequest;
                tRequest = WebRequest.Create("https://fcm.googleapis.com/fcm/send");
                tRequest.Method = "post";
                tRequest.ContentType = " application/json";

                tRequest.Headers.Add(string.Format("Authorization: key={0}", SERVER_API_KEY));
                tRequest.Headers.Add(string.Format("Sender: id={0}", SENDER_ID));

                var a = new
                {
                    notification = new
                    {
                        title,
                        body,
                        icon = "/Image/GIDDINGS_PRIMARY_STACKED_LOGO_DRIFT_RGB.png"//,
                        //click_action,
                        //sound = "mySound"
                    },
                    to = deviceId
                };

                byte[] byteArray = Encoding.UTF8.GetBytes(Newtonsoft.Json.JsonConvert.SerializeObject(a));
                tRequest.ContentLength = byteArray.Length;

                Stream dataStream = tRequest.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();

                WebResponse tResponse = tRequest.GetResponse();
                dataStream = tResponse.GetResponseStream();

                StreamReader tReader = new StreamReader(dataStream);
                string sResponseFromServer = tReader.ReadToEnd();

                tReader.Close();
                dataStream.Close();
                tResponse.Close();

                return sResponseFromServer;
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
    }
}