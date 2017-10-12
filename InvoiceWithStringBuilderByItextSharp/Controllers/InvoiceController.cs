using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;

namespace InvoiceWithStringBuilderByItextSharp.Controllers
{
    public class InvoiceController : Controller
    {
        // GET: Invoice
        public ActionResult Invoice1()
        {

            string clientName = ClientInfo.Name;
            string contactNo = ClientInfo.Contact;
            string email = ClientInfo.Email;
            string address = ClientInfo.Address;

            int clientNo = ClientInfo.ClientID;
            int orderNo = aBillingView.BillNo;
            string sellerBy = paymentInfo.SellerBy;
            string date = ClientInfo.Date.ToString("dd MMMM yyyy");


            double totalCost = paymentInfo.GrandTotal;
            double discount = paymentInfo.Discount;
            double payableAmount = paymentInfo.PayableAmount;
            double due = paymentInfo.Due;
            double advanced = paymentInfo.Advanced;
            string status = paymentInfo.Status;





            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[5] {
                new DataColumn("SL", typeof(string)),
                new DataColumn("Product", typeof(string)),
                new DataColumn("Quantity", typeof(int)),
                new DataColumn("Price", typeof(string)),
                new DataColumn("Net Price", typeof(string))});

            foreach (var aBilling in allBilling)
            {
                dt.Rows.Add(sl, aBilling.Product, aBilling.Quantity, aBilling.Price.ToString("##,###"), aBilling.NetPrice.ToString("##,###"));
                sl++;
                quantityOfBilling += aBilling.Quantity;

            }


            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                {
                    StringBuilder sb = new StringBuilder();

                    //Generate Invoice (Bill) Header.

                    sb.Append("<br/>");
                    sb.Append("<br/>");

                    sb.Append("<table width='100%' cellspacing='0' cellpadding='2' style='font-family: Calibri; font-size: 10pt;'>");

                    sb.Append("<tr><td colspan = '2'></td></tr>");
                    sb.Append("<tr><td style='font-size:14pt;'><b> ");
                    sb.Append(clientName);
                    sb.Append("</b></td><td align = 'right'><b>M M Enterprise </b>");
                    //sb.Append(clientNo.ToString("00000"));
                    sb.Append(" </td></tr>");


                    sb.Append("<tr><td colspan = '2'></td></tr>");
                    sb.Append("<tr><td><b>Date Issued: </b>");
                    sb.Append(date);
                    sb.Append("</td><td align = 'right'>Under Uttara Bank Ltd, Sreemangal Road");
                    //sb.Append(orderNo.ToString("0000"));
                    sb.Append(" </td></tr>");

                    sb.Append("<tr><td colspan = '2'></td></tr>");
                    sb.Append("<tr><td><b>Invoice No: </b>");
                    sb.Append(orderNo.ToString("0000"));
                    sb.Append("</td><td align = 'right'>Moulvibazar, Sylhet");
                    //sb.Append(sellerBy);
                    sb.Append(" </td></tr>");

                    sb.Append("<tr><td colspan = '2'></td></tr>");
                    sb.Append("<tr><td><b>Client-ID: </b>");
                    sb.Append(clientNo.ToString("00000"));
                    sb.Append("</td><td align = 'right'> Seller-By: ");
                    sb.Append(sellerBy);
                    sb.Append(" </td></tr>");

                    sb.Append("</table>");

                    sb.Append("<br />");
                    sb.Append("<br />");

                    //Generate Invoice (Bill) Items Grid.
                    sb.Append("<table border = '0' style='font-family: Calibri; font-size: 10pt;'>");
                    sb.Append("<tr style='background-color: green;font-weight: bold; color:red;'>");
                    foreach (DataColumn column in dt.Columns)
                    {
                        sb.Append("<th>");
                        sb.Append(column.ColumnName);
                        sb.Append("</th>");
                    }
                    sb.Append("</tr>");
                    foreach (DataRow row in dt.Rows)
                    {
                        sb.Append("<tr>");
                        foreach (DataColumn column in dt.Columns)
                        {
                            sb.Append("<td>");
                            sb.Append(row[column]);
                            sb.Append("</td>");
                        }
                        sb.Append("</tr>");
                    }
                    sb.Append("</tr></table>");

                    sb.Append("<br/>");
                    sb.Append("<br/>");
                    sb.Append("<br/>");

                    sb.Append("<table width='100%' cellspacing='0' cellpadding='2' style='font-family: Calibri; font-size: 12pt;'>");
                    sb.Append("<tr><td colspan = '2'></td></tr>");
                    sb.Append("<tr><td><b>Due By </b>");

                    sb.Append("</td><td align = 'right'> Total Due");
                    sb.Append("</td></tr>");

                    sb.Append("<tr><td colspan = '2'></td></tr>");
                    sb.Append("<tr><td>");
                    sb.Append(date);
                    sb.Append("</td><td align = 'right' style='color:red'>");
                    sb.Append(String.Format("{0:N0}", due));
                    sb.Append(" Taka</td></tr>");

                    sb.Append("</table>");



                    sb.Append("<br/>");
                    sb.Append("<br/>");



                    sb.Append("<table width='100%' cellspacing='0' cellpadding='2' style='font-family: Calibri; font-size: 9pt;'>");

                    sb.Append("<tr><td colspan = '2'></td></tr>");
                    sb.Append("<tr><td><span style='font-size:14pt;font-weight:bold; color:red;'>Thank you!</span>");
                    sb.Append("</td><td align = 'right'>01718-283754 | 01919-110496 | 0861-63686");
                    sb.Append("</td></tr>");

                    sb.Append("<tr><td colspan = '2'></td></tr>");
                    sb.Append("<tr><td>          ");
                    sb.Append("</td><td align = 'right'>mmenterprise@gmail.com | mmenterprise.azurewebsites.net");
                    sb.Append("</td></tr>");

                    sb.Append("</table>");

                    //Export HTML String as PDF.
                    StringReader sr = new StringReader(sb.ToString());
                    Document pdfDoc = new Document(PageSize.A4, 40f, 40f, 40f, 0f);
                    HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
                    PdfWriter writer = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
                    pdfDoc.Open();




                    Image png = Image.GetInstance(@"C:\Users\mdsho\Videos\MM Enterprise\icon2.png");
                    pdfDoc.Add(png);




                    htmlparser.Parse(sr);
                    pdfDoc.Close();
                    Response.ContentType = "application/pdf";
                    Response.AddHeader("content-disposition", "attachment;filename=MMEInvoice-" + orderNo + ".pdf");
                    Response.Cache.SetCacheability(HttpCacheability.NoCache);
                    Response.Write(pdfDoc);
                    Response.End();
                }
            }


            return View();
        }
    }
}