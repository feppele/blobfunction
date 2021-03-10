using System;
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using System.Text;  
using System.Security.Cryptography; 
using Microsoft.Azure.Storage.Blob;
//for Powerpoint

using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace Company.Function
{
    public static class BlobTrigger88
    {
        
        [FunctionName("BlobTrigger88")]
        [return: Queue("result-queue", Connection = "blobstore55_STORAGE")]
        public static string Run([BlobTrigger("tutorial-container/{name}", Connection = "blobstore55_STORAGE")]Stream pptxDatei, string name, ILogger log, 
        [Queue("outqueue"),StorageAccount("AzureWebJobsStorage")] ICollector<string> msg,
        [Blob("result-container/{name}", FileAccess.ReadWrite, Connection = "blobstore55_STORAGE")] CloudBlobContainer outputContainer)
        {
            log.LogInformation($"C# Blob trigger ausgelöst durch File\n Name:{name} \n Size: {pptxDatei.Length} Bytes");



            // POWER POINT

            string resultFileText = brandingPolice(pptxDatei);


            // result.txt erstellen & in Blob laden.

            outputContainer.CreateIfNotExistsAsync();
            var blobName = "result.txt";
            var cloudBlockBlob = outputContainer.GetBlockBlobReference(blobName);
            cloudBlockBlob.UploadTextAsync(resultFileText);
            
            string containerURL = Convert.ToString(outputContainer.Uri);
            string resultURL = containerURL + "/" + blobName;

            log.LogInformation(Convert.ToString(outputContainer.Uri));

            //return new OkObjectResult(blobName);



            // URL von result.txt in Queue
            return resultURL;
        }





        public static string brandingPolice(Stream pptxDatei){

            string[] textArray;

            textArray = presentationToTextArray(pptxDatei);

            // int i=1;
            // foreach(string text in textArray){
                
            //     log.LogInformation("Folie" + Convert.ToString(i) + ":   "+ text);
            //     i++;
            // }

            string[] falscheMarken = {"Windows Azure","Windows 365","Windows Office"};


            string[] resultArray = new string[presentationSlidesCount(pptxDatei)];

            int folienNr =1;
            foreach( string text in textArray){

                for(int k=0; k< falscheMarken.Count(); ++k){

                    if(text.Contains(falscheMarken[k])){
                        
                        resultArray[folienNr-1] += falscheMarken[k] + "   ";
                        
                    }
                }
    
                folienNr++;

            }

            string resultFileText = new string("Branding Violations: \n\n");
            folienNr=1;

            foreach(string text in resultArray){
                
                if(text != ""){
                    
                    resultFileText += "Folie " + Convert.ToString(folienNr) + ": " + text + "\n";

                }

                folienNr++;

            }

            return resultFileText;
        }

        public static int presentationSlidesCount(Stream pptxDatei){
            PresentationDocument presentationDocument = PresentationDocument.Open(pptxDatei, false);


            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            int slidesCount = 0;
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            slidesCount = presentationPart.SlideParts.Count();

            return slidesCount;
        }

        public static string[] presentationToTextArray(Stream pptxDatei){
            
            PresentationDocument presentationDocument = PresentationDocument.Open(pptxDatei, false);

            // Check for a null document object.

            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            int slidesCount = 0;
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            slidesCount = presentationPart.SlideParts.Count();
            

            string[] textArray = new string[slidesCount];

            string slideText;
            for (int i = 0; i < slidesCount; i++)
            {

                OpenXmlElementList slideIds = presentationPart.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[i] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart) presentationPart.GetPartById(relId);

                // Build a StringBuilder object.
                StringBuilder paragraphText = new StringBuilder();

                // Get the inner text of the slide:
                IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
                foreach (A.Text text in texts)
                {
                    paragraphText.Append(text.Text);
                }

                slideText = paragraphText.ToString();

                textArray[i]=slideText;

                //log.LogInformation("Slide #{0} contains: {1}", i + 1, slideText);

            }



            return textArray;
        }



    }





}






        // [FunctionName("BlobTrigger88")]
        // [return: Queue("myqueue-items", Connection = "blobstore55_STORAGE")]
        // public static string Run([BlobTrigger("tutorial-container/{name}", Connection = "blobstore55_STORAGE")]Stream myBlob, string name, ILogger log, 
        // [Queue("outqueue"),StorageAccount("AzureWebJobsStorage")] ICollector<string> msg,
        // [Blob("text-container/{name}", FileAccess.Write, Connection = "blobstore55_STORAGE")] Stream txtOutput)
        // {
        //     log.LogInformation($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");

        //     byte[] buf = new byte[myBlob.Length];
        //     //Convert.ToInt32(myBlob.Length)
        //     myBlob.Read(buf,0,Convert.ToInt32(myBlob.Length));
        //     string s = myBlob.ToString();

        //     log.LogInformation(BitConverter.ToString(buf));
        //     log.LogInformation( Encoding.UTF8.GetString(buf,0,buf.Length) );

        //     byte[] buf2 = new byte[5];
        //     buf2[0]= Convert.ToByte("h");
        //     buf2[1]= Convert.ToByte("i");



        //     string txtPath = "./data/result.txt";
        //     File.WriteAllTextAsync(txtPath, "Beispiel Text...");

        //     txtOutput.Write()
        //     txtOutput = File.OpenRead(txtPath);

            