using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;
using static Sistema_Indicadores.Clases.FCM.FCM;

namespace Sistema_Indicadores.Clases
{
    public class Notificaciones
    {



        //public void SendNotification(string title, string body)
        //{
        //    try
        //    {
        //        WebRequest tRequest = WebRequest.Create("https://fcm.googleapis.com/fcm/send");
        //        tRequest.Method = "post";

        //        string SERVER_KEY_TOKEN = "AIzaSyANuakkocCPEoISIJG7DID6StVrEu3Dm80";
        //        string deviceId = "bda93f56ce507588";
        //        var SENDER_ID = "G-3HCJ4PZ3WW";

        //        //serverKey - Key from Firebase cloud messaging server  
        //        tRequest.Headers.Add(string.Format("Authorization: key={0}", SERVER_KEY_TOKEN));
        //        //Sender Id - From firebase project setting  
        //        tRequest.Headers.Add(string.Format("Sender: id={0}", SENDER_ID));
        //        tRequest.ContentType = "application/json";
        //        var payload = new
        //        {
        //            to = deviceId,
        //            priority = "high",
        //            content_available = true,
        //            notification = new
        //            {
        //                body,
        //                title,
        //                badge = 2,
        //                icon = "/Image/GIDDINGS_PRIMARY_STACKED_LOGO_DRIFT_RGB.png"
        //            },
        //            data = new
        //            {
        //                codigo = deviceId
        //            }
        //        };

        //        string postbody = JsonConvert.SerializeObject(payload).ToString();
        //        Byte[] byteArray = Encoding.UTF8.GetBytes(postbody);
        //        tRequest.ContentLength = byteArray.Length;
        //        using (Stream dataStream = tRequest.GetRequestStream())
        //        {
        //            dataStream.Write(byteArray, 0, byteArray.Length);
        //            using (WebResponse tResponse = tRequest.GetResponse())
        //            {
        //                using (Stream dataStreamResponse = tResponse.GetResponseStream())
        //                {
        //                    if (dataStreamResponse != null) using (StreamReader tReader = new StreamReader(dataStreamResponse))
        //                        {
        //                            String sResponseFromServer = tReader.ReadToEnd();
        //                            //result.Response = sResponseFromServer;
        //                        }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception e) {
        //        e.ToString();
        //    }
        //}

        public string SendNotificationJSON(string title, string body)
        {
            try
            {
                string SERVER_KEY_TOKEN = "AIzaSyANuakkocCPEoISIJG7DID6StVrEu3Dm80";
                string deviceId = "bda93f56ce507588";
                var SENDER_ID = "G-3HCJ4PZ3WW";

                WebRequest tRequest;
                tRequest = WebRequest.Create("https://fcm.googleapis.com/fcm/send");
                tRequest.Method = "post";
                tRequest.ContentType = " application/json";

                tRequest.Headers.Add(string.Format("Authorization: key={0}", SERVER_KEY_TOKEN));
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