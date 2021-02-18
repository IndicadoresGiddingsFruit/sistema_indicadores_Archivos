using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace Sistema_Indicadores.Controllers
{
    public class FCMController : Controller
    {
        //[System.Web.Http.HttpGet]
        //[System.Web.Http.Route("SendMessage")]
        //public IHttpActionResult SendMessage()
        //{
        //    var data = new
        //    {
        //        to = "bda93f56ce507588",
        //        data = new
        //        {
        //            message = "",
        //            name = "Shoury",
        //            userId = "1",
        //            status = true
        //        }
        //    };
        //    SendNotification(data);
        //    return Ok();
        //}

        //public void SendNotification(object data)
        //{
        //    var serializer = new JavaScriptSerializer();
        //    var json = serializer.Serialize(data);
        //    Byte[] byteArray = Encoding.UTF8.GetBytes(json);

        //    SendNotification(byteArray);
        //}

        //public void SendNotification(Byte[] byteArray)
        //{
        //    try
        //    {
        //        string server_api_key = ConfigurationManager.AppSettings["SERVER_API_KEY"];
        //        string sender_id = ConfigurationManager.AppSettings["SENDER_ID"];

        //        WebRequest tRequest = WebRequest.Create("https://fcm.googleapis.com/fcm/send");
        //        tRequest.Method = "post";
        //        tRequest.ContentType = " application/json";
        //        tRequest.Headers.Add($"Authorization: key={server_api_key}");
        //        tRequest.Headers.Add($"Sender: id={sender_id}");

        //        tRequest.ContentLength = byteArray.Length;
        //        Stream dataStream = tRequest.GetRequestStream();
        //        dataStream.Write(byteArray, 0, byteArray.Length);
        //        dataStream.Close();

        //        WebResponse tResponse = tRequest.GetResponse();
        //        dataStream = tResponse.GetResponseStream();
        //        StreamReader tReader = new StreamReader(dataStream);

        //        string sResponseFromServer = tReader.ReadToEnd();

        //        tReader.Close();
        //        dataStream.Close();
        //        tResponse.Close();
        //    }
        //    catch (Exception e) { e.ToString(); }
        //}
    }
}
