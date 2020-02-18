using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace estagioprobatorioback.Comum
{
    public class AlfrescoService
    {
        private static readonly HttpClient client = new HttpClient();

        public AlfrescoService()
        {
            var byteArray = Encoding.ASCII.GetBytes("credentials");
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
        }
        private string nodeRef;
        public void UploadDocumento(Stream sFile, string nameFile, int idAvaliacao, bool overwrite)
        {
            try
            {
                using (var formData = new MultipartFormDataContent())
                {
                    formData.Add(new StreamContent(sFile), "filedata", idAvaliacao + "-IDENT" + "-" + nameFile);
                    formData.Add(new StringContent("STRUCTURE"), "id");
                    formData.Add(new StringContent("documentLibrary"), "containerid");
                    formData.Add(new StringContent("/FOLDER"), "uploaddirectory");
                    formData.Add(new StringContent("cm:content"), "contenttype");
                    formData.Add(new StringContent(overwrite.ToString()), "overwrite");
                    formData.Add(new StringContent(""), "description");

                    var response = client.PostAsync("http://alfrescourl", formData).Result;

                    dynamic jsonResult = null;
                    if (response.Content != null)
                    {
                        jsonResult = JObject.Parse(response.Content.ReadAsStringAsync().Result);
                        nodeRef = jsonResult["nodeRef"].ToString();
                    }
                    else
                    {
                        throw new Exception(response.RequestMessage.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public Stream  DownloadDocumento(string nodeRef)
        {
            var formatNodRef = nodeRef.Replace("workspace://SpacesStore/", "");
            var response = client.GetStreamAsync("alfresco-download-url" + formatNodRef).Result;
            return response;
        }

        public string GetNodeRef()
        {
            return this.nodeRef;
        }
    }
}