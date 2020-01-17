using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using System.Net.Mail;

namespace H2oReport
{

   namespace H2oDiagnosis2.Models
   {
      public enum ApartmentType
      {
         IndividualHouse,
         Apartment,
         SemiGated,
         Gated
      };

      public enum RWHSType
      {
         Storage,// sampu
         Bores,// 
         Pits
      };

      public enum RecPlantType
      {
         Non_PotableUsage,// 
         DomesticUsage
      };

      public struct MyLocation
      {
         public float m_Lat;
         public float m_Long;
         public float m_Alt;
      };

      public class H20DiagnosticsInputData
      {
         //public string m_PageName { get; set; }
         public string m_Name { get; set; }
         public long m_MobileNumber { get; set; }
         public string m_EmailId { get; set; }// Optional
         public int m_CANID { get; set; }// Optional
         public MyLocation m_Location { get; set; }

         public int m_AptType { get; set; }
         public double m_RoofArea { get; set; }
         public int m_FlatCount { get; set; }
         public int m_PeopleCount { get; set; }// Optional
         public bool m_WaterMeters { get; set; }// Optional
         public bool m_TapWaterSavers { get; set; }// Optional

         public bool m_RWHSExists { get; set; }
         public int m_RWHSType { get; set; }// // Optional based on above
         public bool m_RWHSIsOverFlow { get; set; }
         public int m_BoreWellCount { get; set; }
         public int m_FunctBoreWellCount { get; set; }

         public bool m_RecPlantExists { get; set; }
         public int m_RecPlantType { get; set; }// Optional
         public double m_RecPlantCapacity { get; set; }// Liters
      };

      public struct H20DiagnosticsOutputData
      {
         public double m_UsageLtMin;
         public double m_UsageLtMax;
         public double m_RWHSWaterLtMin;
         public double m_RWHSWaterLtMax;
         public double m_RWHSCostMin;
         public double m_RWHSCostMax;
         public double m_RecPlantWaterLtMin;
         public double m_RecPlantWaterLtMax;
         public double m_RecPlantCostMin;
         public double m_RecPlantCostMax;
         public double m_PowerBill;
         public double m_UsageScore;
      };

      public struct H20DiagnosticsData
      {
         public H20DiagnosticsInputData m_H20DiagIpData;
         public H20DiagnosticsOutputData m_H20DiagOpData;
      };

      public class H2oReporter
      {

         public void Report(H20DiagnosticsData m_h2oDiagsData)
         {
            try
            {
               var winword = new Microsoft.Office.Interop.Word.Application();

               //Set animation status for word application  
               winword.ShowAnimation = false;

               //Set status for word application is to be visible or not.  
               winword.Visible = false;

               //Create a missing variable for missing value  
               object missing = System.Reflection.Missing.Value;

               //Create a new document  
               Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

               //Add header into the document  
               foreach (Section section in document.Sections)
               {
                  //Get the header range and add the header details.  
                  Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                  headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                  headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                  headerRange.Font.ColorIndex = WdColorIndex.wdBlue;
                  headerRange.Font.Size = 20;
                  headerRange.Font.Bold = 1;
                  headerRange.Text = "H2O Diagnostic Report";
               }

               object styleNormal = "Normal";

               //Add paragraph with Heading 1 style  
               string line1 = "ID: " + m_h2oDiagsData.m_H20DiagIpData.m_CANID.ToString() + "\t\t\t\t\t\t" + "Mobile Number: " + m_h2oDiagsData.m_H20DiagIpData.m_MobileNumber;
               Paragraph paraline1 = document.Content.Paragraphs.Add(ref missing);
               paraline1.Range.set_Style(ref styleNormal);
               paraline1.Range.Font.Size = 10;
               paraline1.Range.Font.Bold = 1;
               paraline1.Range.Text = line1;
               paraline1.Range.InsertParagraphAfter();

               string line2 = "Name: " + m_h2oDiagsData.m_H20DiagIpData.m_Name + "\t\t\t\t\t\t" + "Email: " + m_h2oDiagsData.m_H20DiagIpData.m_EmailId;
               Paragraph paraline2 = document.Content.Paragraphs.Add(ref missing);
               paraline2.Range.set_Style(ref styleNormal);
               paraline2.Range.Font.Size = 10;
               paraline2.Range.Font.Bold = 1;
               paraline2.Range.Text = line2;
               paraline2.Range.InsertParagraphAfter();

               string line3 = "Your Usage Score is: ";
               Paragraph paraline3 = document.Content.Paragraphs.Add(ref missing);
               paraline3.Range.set_Style(ref styleNormal);
               paraline3.Range.Font.Size = 10;
               paraline3.Range.Font.Bold = 1;
               paraline3.Range.Text = line3;
               paraline3.Range.InsertParagraphAfter();

               double Usage = m_h2oDiagsData.m_H20DiagOpData.m_UsageScore;
               Shape usage = document.Shapes.AddShape(5, 175, 155, 200, 20);
               if (Usage < 35)
                  usage.Fill.ForeColor.RGB = 16711680;
               else if (Usage >= 35 && Usage < 50)
                  usage.Fill.ForeColor.RGB = 12632256;
               else if (Usage >= 50 && Usage < 75)
                  usage.Fill.ForeColor.RGB = 65280;
               else if (Usage >= 75 && Usage < 100)
                  usage.Fill.ForeColor.RGB = 10082794;

               //string line13 = "Your Usage Score is: ";
               Paragraph paraline13 = document.Content.Paragraphs.Add(ref missing);
               paraline13.Range.set_Style(ref styleNormal);
               paraline13.Range.Font.Size = 16;
               paraline13.Range.Font.Bold = 1;
               paraline13.Range.Text = "\t\t\t\t\t" + Usage.ToString();
               paraline13.Range.InsertParagraphAfter();

               string line4 = "\nYour Current Usage is: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_UsageLtMin.ToString() + " litres" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_UsageLtMax.ToString() + " litres";
               Paragraph paraline4 = document.Content.Paragraphs.Add(ref missing);
               paraline4.Range.set_Style(ref styleNormal);
               paraline4.Range.Font.Size = 10;
               paraline4.Range.Font.Bold = 0;
               paraline4.Range.Text = line4;
               paraline4.Range.InsertParagraphAfter();

               string sline5 = "";
               if (m_h2oDiagsData.m_H20DiagIpData.m_RWHSExists)
               {
                  if (m_h2oDiagsData.m_H20DiagIpData.m_RWHSIsOverFlow)
                  {
                     sline5 = "\nGood that you have installed Rainwater harvesting pits. If your pits are overflowing it means water is not sent to ground at all, Clean the pit immediately. The best practice is to directly collect and store rainwater in your regular sump and excess water can be redirected to either your borewell (functional or de-functional (de-functional bore is a bore which used to give water but not anymore)) or recharging pits. With this you can ensure 100 % usage of your rain water." +
                              "\nYou are almost losing: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RWHSWaterLtMin.ToString() + " litres" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantWaterLtMax.ToString() + " litres" +
                              "\nYou are almost losing: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RWHSCostMin.ToString() + " rupees" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RWHSCostMax.ToString() + " rupees";
                  }
                  else
                  {
                     //if (overflow)
                     {
                        string sline = "";
                        if (m_h2oDiagsData.m_H20DiagIpData.m_RWHSType == 0) // Storage
                           sline = "\nGood that you are storing rainwater, reusing and redirecting excess water to either borewell (functional or de-functional (de-functional bore is a bore which used to give water but not anymore)) or recharging pits. This is the best practice.";
                        else if (m_h2oDiagsData.m_H20DiagIpData.m_RWHSType == 1) // Directly sent to Recharge Pits
                           sline = "\nGood that you have installed Rain water harvesting pits. The best practice is to directly collect and store rainwater in your regular sump and excess water can be redirected to either your borewell (functional or de-functional (de-functional bore is a bore which used to give water but not anymore)) or recharging pits. With this you can ensure 100% usage of your rain water.";
                        else if (m_h2oDiagsData.m_H20DiagIpData.m_RWHSType == 2) // Directly sent to borewell
                           sline = "\nGood that you have are redirecting your collected rain water to your borewell. The best practice is to directly collect and store rainwater in your regular sump and excess water can be redirected to either your borewell (functional or de-functional (de-functional bore is a bore which used to give water but not anymore)) or recharging pits. With this you can ensure 100% usage of your rain water.";

                        Paragraph line = document.Content.Paragraphs.Add(ref missing);
                        line.Range.set_Style(ref styleNormal);
                        line.Range.Font.Size = 10;
                        line.Range.Font.Bold = 0;
                        line.Range.Text = sline;
                        line.Range.InsertParagraphAfter();

                        sline5 = "\nYou are almost saving: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RWHSWaterLtMin.ToString() + " litres" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RWHSWaterLtMax.ToString() + " litres" +
                                 "\nYou are almost saving: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RWHSCostMin.ToString() + " rupees" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RWHSCostMax.ToString() + " rupees" +
                                 "\nRemember to clean your Rainwater Harvesting pits before rainy season every year.\nAlso remember that rain water is the purest form of water.";
                     }
                  }
               }
               else
               {
                  sline5 = "Installing Rain water harvesting pits could benefit you with some water availability.\nRemember that rain water is the purest form of water." +
                           "\nYou are almost losing: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RWHSWaterLtMin.ToString() + " litres" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantWaterLtMax.ToString() + " litres" +
                           "\nYou are almost losing: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RWHSCostMin.ToString() + " rupees" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RWHSCostMax.ToString() + " rupees" +
                           "\nThe best practice is to directly collect and store rainwater in your regular sump and excess water can be redirected to either your borewell (functional or de-functional (de-functional bore is a bore which used to give water but not anymore)) or recharging pits. With this you can ensure 100 % usage of your rain water.";
               }

               Paragraph paraline5 = document.Content.Paragraphs.Add(ref missing);
               paraline4.Range.set_Style(ref styleNormal);
               paraline4.Range.Font.Size = 10;
               paraline4.Range.Font.Bold = 0;
               paraline4.Range.Text = sline5;
               paraline4.Range.InsertParagraphAfter();

               string sline6 = "";
               if (m_h2oDiagsData.m_H20DiagIpData.m_RecPlantExists)
               {
                  if (m_h2oDiagsData.m_H20DiagIpData.m_RecPlantType == 0) // non-potable
                  {
                     sline6 = "Good that you are recycling water and reusing them for various purposes. It reduces environmental pollution from improper wastewater disposal. You may think of installing a recycling plant which can be used for domestic purpose." +
                              "\nYou are almost saving: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantWaterLtMin.ToString() + " litres" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantWaterLtMax.ToString() + " litres" +
                              "\nYou are almost saving: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantCostMin.ToString() + " rupees" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantCostMax.ToString() + " rupees";
                  }
                  else // domestic use
                  {
                     sline6 = "Good that you are recycling water and reusing them for various purposes. It reduces environmental pollution from improper wastewater disposal." +
                              "\nYou are almost saving: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantWaterLtMin.ToString() + " litres" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantWaterLtMax.ToString() + " litres" +
                              "\nYou are almost saving: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantCostMin.ToString() + " rupees" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantCostMax.ToString() + " rupees";
                  }
               }
               else
               {
                  sline6 = "Installing a water recycling plant would help you to save almost 30%-50% of your water usage. Recycled water may be reused for irrigation of gardens and agricultural fields or replenishing surface water and groundwater. Reused water may also be used for drinking purposes. Although it may incur some cost initially it will benefit you in long run. The cost may increase depending on the quality of output water." +
                           "\nYou are almost losing: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantWaterLtMin.ToString() + " litres" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantWaterLtMax.ToString() + " litres" +
                           "\nYou are almost losing: \n" + "Min: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantCostMin.ToString() + " rupees" + "\t\t\t\t\t\t" + "Max: " + m_h2oDiagsData.m_H20DiagOpData.m_RecPlantCostMax.ToString() + " rupees";
               }

               Paragraph paraline6 = document.Content.Paragraphs.Add(ref missing);
               paraline6.Range.set_Style(ref styleNormal);
               paraline6.Range.Font.Size = 10;
               paraline6.Range.Font.Bold = 0;
               paraline6.Range.Text = sline6;
               paraline6.Range.InsertParagraphAfter();

               //Save the document  
               object filename = @"c:\H2ODiagnostics.docx";
               document.SaveAs2(ref filename);
               document.Close(ref missing, ref missing, ref missing);
               document = null;

               winword.Quit(ref missing, ref missing, ref missing);
               winword = null;
               //MessageBox.Show("Document created successfully !");

               System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
               SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
               mail.From = new MailAddress("H2ODiagnostics@gmail.com");
               mail.To.Add("hepsibasushma@gmail.com");
               mail.Subject = "H2ODiagnostics Report";
               mail.Body = "Attached is your report";

               System.Net.Mail.Attachment attachment;
               attachment = new System.Net.Mail.Attachment(@"c:\H2ODiagnostics.docx");
               mail.Attachments.Add(attachment);

               SmtpServer.Port = 587;
               SmtpServer.Credentials = new System.Net.NetworkCredential("h20diagnostics@gmail.com", "Hexathon");
               SmtpServer.EnableSsl = true;

               SmtpServer.Send(mail);
            }
            catch (Exception ex)
            {
               //DisplayAlert(ex.Message);
            }
         }
      }
   }
}
