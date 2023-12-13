using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Security.Claims;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using RACE.Models;
using Microsoft.AspNetCore.Http;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;

namespace RACE.Controllers
{
    [Authorize]
    public class PaymentCustController : Controller
    {
        private readonly RACEDbContext _context;
        IDataProtector dataProtector;

        public PaymentCustController(RACEDbContext context, IDataProtectionProvider provider)
        {
            _context = context;
            dataProtector = provider.CreateProtector("ReDAntSols");
        }


        // GET: Payment Voucher
        public ActionResult SalesInvoice_Index()
        {

            PrepareSalesInvoiceIndexView();
            return View(new List<SalesInvoice>());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SalesInvoice_Index(string CustomerID, string InvoiceNo, string ReferenceNo)
        {
            var claimSite = User.FindFirstValue("SiteID");

            var payments = await _context.SalesInvoice
                .Where(e => e.SiteID == claimSite
                && e.CustomerID == (CustomerID != null ? CustomerID : e.CustomerID)
                && e.InvoiceNo == (InvoiceNo != null ? InvoiceNo : e.InvoiceNo)
                //&& e.ReferenceNo == (ReferenceNo != null ? ReferenceNo : e.ReferenceNo) // not working ???
                && (ReferenceNo != null ? e.ReferenceNo == ReferenceNo : true)
                ).ToListAsync();


            foreach (var pay in payments)
            {
                pay.Customer = _context.Customer.FirstOrDefault(e => e.CustomerID == pay.CustomerID);
                pay.ProtectedInvoiceNo = dataProtector.Protect(pay.InvoiceNo);
            }
            PrepareSalesInvoiceIndexView();
            return View(payments);
        }

        private void PrepareSalesInvoiceIndexView()
        {
            var claimSite = User.FindFirstValue("SiteID");
            List<SelectListItem> MenuList = new List<SelectListItem>();
            var psSQL = from s in _context.Customer
                            //join st in _context.Site on s.SiteID equals st.SiteID into temp
                            //from st2 in temp.DefaultIfEmpty()
                            //where s.SiteID == (claimSite != "Admin" ? claimSite : s.SiteID) || s.SiteID == null
                        where s.SiteID == claimSite || s.SiteID == null
                        orderby s.CustomerCode ascending
                        select s;
            var psList = psSQL.ToList();
            if (psList.Count > 0)
            {
                for (int i = 0; i < psList.Count; i++)
                {
                    SelectListItem item = new SelectListItem { Value = psList[i].CustomerID, Text = psList[i].CustomerCode + " - " + psList[i].CustomerName };
                    MenuList.Add(item);
                }
            }
            ViewBag.CustomerList = MenuList;

        }

        // GET: Payment Voucher/Create
        public IActionResult SalesInvoice_Create()
        {

            var claimSite = User.FindFirstValue("SiteID");


            List<SelectListItem> MenuList = new List<SelectListItem>();
            var psSQL = from s in _context.Customer
                        where s.SiteID == claimSite || s.SiteID == null
                        orderby s.CustomerCode ascending
                        select s;
            var psList = psSQL.ToList();
            if (psList.Count > 0)
            {
                for (int i = 0; i < psList.Count; i++)
                {
                    SelectListItem item = new SelectListItem { Value = psList[i].CustomerID, Text = psList[i].CustomerCode + " - " + psList[i].CustomerName };
                    MenuList.Add(item);
                }
            }
            ViewBag.CustomerList = MenuList;


            MenuList = new List<SelectListItem>();
            var psList2 = (from s in _context.Product
                           where (s.SiteID == claimSite || s.SiteID == null)
                           && s.LoanFlag == false  //exclude loan item
                           && s.ItemLoanFlag == false //exclude good loan item
                           orderby s.ProductCode ascending
                           select s).ToList();
            if (psList2.Count > 0)
            {
                for (int i = 0; i < psList2.Count; i++)
                {
                    SelectListItem item = new SelectListItem { Value = psList2[i].ProductID, Text = psList2[i].ProductCode };
                    MenuList.Add(item);
                }
            }
            ViewBag.ProductList = MenuList;


            return View();
        }

        // POST: Sales Invoice/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SalesInvoice_Create([Bind("CustomerID,ReferenceNo,TotalAmount,Remark")] SalesInvoice invoice, List<string> Description, List<string> ProductID, List<string> refID, List<string> InvoiceType, List<string> UOMId, List<decimal> UnitPrice, List<decimal> Qty, string Action, string CustomerID, List<IFormFile> files, string InvoiceDate)
        {

            if (ModelState.IsValid)
            {

                List<string> DOList = new List<String>();

                var Customer = _context.Customer.SingleOrDefault(e => e.CustomerID == CustomerID);

                IDSetupController id = new IDSetupController(_context);
                invoice.InvoiceNo = id.InternalGetNewID("SalesInvoice");
                invoice.InvoiceDate = DateTime.ParseExact(InvoiceDate, "dd/MM/yyyy", null);
                invoice.CreatedBy = User.Identity.Name;
                invoice.CreatedDate = DateTime.Now;
                invoice.SiteID = User.FindFirstValue("SiteID");

                if (Action == "Draft")
                {
                    invoice.InvoiceStatus = "Draft";
                    //invoice.InvoiceStage = "invoice Voucher";
                }
                else if (Action == "Submit")
                {
                    //invoice.InvoiceStatus = "Pending";
                    //invoice.InvoiceStage = "invoice";
                    invoice.InvoiceStatus = "Payment";
                }

                invoice.PaymentTermsValue = Customer.PaymentTermsValue;
                invoice.PaymentTermsType = Customer.PaymentTermsType;
                invoice.BillingAddress1 = Customer.BillingAddress1;
                invoice.BillingAddress2 = Customer.BillingAddress2;
                invoice.BillingCity = Customer.BillingCity;
                invoice.BillingState = Customer.BillingState;
                invoice.BillingPostcode = Customer.BillingPostcode;
                invoice.BillingCountry = Customer.BillingCountry;
                invoice.BillingPhone = Customer.BillingPhone;
                invoice.BillingFax = Customer.BillingFax;
                invoice.BillingPIC = Customer.BillingPIC;

                //var pRefId = "";
                for (int i = 0; i < ProductID.Count; i++)
                {
                    if (ProductID[i] != null && ProductID[i] != "")
                    {
                        //if (refID[i] == "" || refID[i] == null)
                        //{
                        //    pRefId = "NA";
                        //}else
                        //{
                        //    pRefId = refID[i];
                        //}
                        SalesInvoiceDetail invoiceDetail = new SalesInvoiceDetail();
                        invoiceDetail.InvoiceNo = invoice.InvoiceNo;
                        invoiceDetail.RefID = refID[i];
                        invoiceDetail.Description = Description[i];
                        //invoiceDetail.ProductCode = ProductCode[i];
                        invoiceDetail.ProductID = ProductID[i];
                        invoiceDetail.InvoiceType = InvoiceType[i];
                        invoiceDetail.UOM = UOMId[i];
                        invoiceDetail.UnitPrice = UnitPrice[i];
                        invoiceDetail.Qty = Qty[i];
                        invoiceDetail.SubTotal = decimal.Round(UnitPrice[i] * Qty[i], 2, MidpointRounding.AwayFromZero);
                        invoice.TotalAmount += decimal.Round(invoiceDetail.SubTotal, 2, MidpointRounding.AwayFromZero);
                        _context.Add(invoiceDetail);

                        if (InvoiceType[i] == "Ticket")
                        {
                            var ticketDetail = _context.TicketDetail.FirstOrDefault(e => e.TicketID == refID[i] && e.ProductID == ProductID[i]);
                            ticketDetail.SINo = invoice.InvoiceNo;
                            ticketDetail.SIFlag = true;
                            _context.Update(ticketDetail);
                        }
                        //else if (InvoiceType[i] == "SalesOrder")  //change to DO
                        //{
                        //    var SalesOrderDetail = _context.SalesOrderDetail.FirstOrDefault(e => e.SalesOrderNo == refID[i] && e.ProductID == ProductID[i]);
                        //    SalesOrderDetail.SINo = invoice.InvoiceNo;
                        //    SalesOrderDetail.SIFlag = true;
                        //    _context.Update(SalesOrderDetail);
                        //}
                        else if (InvoiceType[i] == "DO")
                        {
                            var DeliveryOrderDetail = _context.DeliveryOrderDetail.FirstOrDefault(e => e.DeliveryOrderNo == refID[i] && e.ProductID == ProductID[i]);
                            DeliveryOrderDetail.SINo = invoice.InvoiceNo;
                            DeliveryOrderDetail.SIFlag = true;
                            _context.Update(DeliveryOrderDetail);
                            DOList.Add(refID[i]);
                        }
                    }
                }
                invoice.TotalInWord = Common.ConvertToWords(invoice.TotalAmount.ToString());

                foreach (var file in files)
                {
                    if (!(file == null || file.Length == 0))
                    {
                        string[] temp = file.FileName.Split(".");
                        string extension = temp[temp.Length - 1];

                        var path = Path.Combine(_context.AppSettings.SingleOrDefaultAsync(a => a.Name == "SalesInvoiceUploadFolder").Result.Value + "\\\\" + invoice.InvoiceNo + "\\\\");
                        string directoryName = Path.GetDirectoryName(path);

                        if (!System.IO.Directory.Exists(path))
                            System.IO.Directory.CreateDirectory(path);

                        path += invoice.InvoiceNo + "_" + (Directory.GetFiles(directoryName).Length + 1).ToString() + "." + extension;

                        using (var stream = new FileStream(path, FileMode.Create))
                        {
                            await file.CopyToAsync(stream);
                            SalesInvoiceDocument document = new SalesInvoiceDocument();
                            document.InvoiceNo = invoice.InvoiceNo;
                            document.FileName = invoice.InvoiceNo + "_" + (Directory.GetFiles(directoryName).Length).ToString() + "." + extension;
                            document.Status = "Active";
                            document.CreatedBy = User.Identity.Name;
                            document.CreatedDate = DateTime.Now;
                            _context.Add(document);
                        }
                    }
                }

                _context.Add(invoice);



                _context.Add(
                    new Log()
                    {
                        EmployeeID = User.Identity.Name,
                        SiteID = User.FindFirstValue("SiteID"),
                        Controller = ControllerContext.ActionDescriptor.ControllerName,
                        Method = ControllerContext.ActionDescriptor.ActionName,
                        Action = Action,
                        ReferenceID = invoice.InvoiceNo,
                        ActionDate = DateTime.Now
                    }
                );

                await _context.SaveChangesAsync();


                //check if all DODetails issued SI then change status to Closed
                List<string> DODistinct = new List<String>();
                DODistinct.AddRange(DOList.Distinct());

                for (int i = 0; i < DODistinct.Count; i++)
                {
                    if (DODistinct[i] != null && DODistinct[i] != "")
                    {
                        var Result = _context.DeliveryOrderDetail.Any(e => e.SIFlag == false && e.DeliveryOrderNo == DODistinct[i]);
                        if (Result == false)
                        {
                            var DO = _context.DeliveryOrder.FirstOrDefault(e => e.DeliveryOrderNo == DODistinct[i]);
                            DO.DeliveryOrderStatus = "Closed";
                            _context.Update(DO);
                        }
                    }

                }
                await _context.SaveChangesAsync();


                if (Action == "Draft")
                {
                    TempData["Notice"] = "Sales Invoice (" + invoice.InvoiceNo + ") has been saved.";
                    return RedirectToAction(nameof(SalesInvoice_Create));
                }
                else if (Action == "Submit")
                {
                    TempData["Success"] = "Sales Invoice (" + invoice.InvoiceNo + ") has been created successfully.";
                    return RedirectToAction(nameof(SalesInvoice_Details), new { id = dataProtector.Protect(invoice.InvoiceNo) });
                }
            }
            else
            {
                var errors = string.Join(" | ", ModelState.Values
                            .SelectMany(v => v.Errors)
                            .Select(e => e.ErrorMessage));
                TempData["Error"] = "Invalid ModelState" + errors;

            }
            return RedirectToAction(nameof(SalesInvoice_Create));

        }



        //// GET: Sales Invoice/Edit
        public async Task<IActionResult> SalesInvoice_Edit(string id)
        {

            var invoice = await _context.SalesInvoice
                .Include(p => p.Customer).Include(e => e.SalesInvoiceDetail)
                .SingleOrDefaultAsync(m => m.InvoiceNo == dataProtector.Unprotect(id));

            if (invoice == null)
            {
                return NotFound();
            }

            invoice.SalesInvoiceDocument = _context.SalesInvoiceDocument.Where(e => e.InvoiceNo == invoice.InvoiceNo && e.Status == "Active").ToList();

            foreach (var detail in invoice.SalesInvoiceDetail)
            {
                detail.Product = _context.Product.FirstOrDefault(e => e.ProductID == detail.ProductID);
            }

            ViewBag.DocumentURL = _context.AppSettings.SingleOrDefaultAsync(a => a.Name == "SalesInvoiceDisplayFolder").Result.Value;

            return View(invoice);
        }

        // POST: Sales Invoice/Edit
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SalesInvoice_Edit([Bind("InvoiceNo")] SalesInvoice invoice, List<string> Description, List<string> ProductID, List<string> refID, List<string> InvoiceType, List<decimal> UnitPrice, List<decimal> Qty, string Action, List<string> CurrentFile, List<IFormFile> files)
        {
            if (ModelState.IsValid)
            {
                SalesInvoice SalesInvoice = _context.SalesInvoice.Include(e => e.SalesInvoiceDetail).FirstOrDefault(e => e.InvoiceNo == invoice.InvoiceNo && e.InvoiceStatus == "Draft");
                SalesInvoice.ModifiedBy = User.Identity.Name;
                SalesInvoice.ModifiedDate = DateTime.Now;

                if (Action == "Draft")
                {
                    SalesInvoice.InvoiceStatus = "Draft";
                    //SalesInvoice.InvoiceStage = "invoice Voucher";
                }
                else if (Action == "Cancel")
                {
                    SalesInvoice.InvoiceStatus = "Canceled";
                }
                else if (Action == "Submit")
                {
                    //SalesInvoice.InvoiceStatus = "Pending";
                    //SalesInvoice.InvoiceStage = "invoice";
                    SalesInvoice.InvoiceStatus = "Payment";
                }

                if (Action == "Draft" || Action == "Submit")
                {
                    SalesInvoice.TotalAmount = 0;
                    if (SalesInvoice.SalesInvoiceDetail.Count > 0)
                    {
                        for (int i = SalesInvoice.SalesInvoiceDetail.Count - 1; i >= 0; i--)
                        {
                            if (!ProductID.Contains(SalesInvoice.SalesInvoiceDetail[i].ProductID))
                            {
                                _context.Remove(SalesInvoice.SalesInvoiceDetail[i]);
                            }
                        }

                    }

                    var CurrentFileList = _context.SalesInvoiceDocument.Where(e => e.InvoiceNo == invoice.InvoiceNo && e.Status == "Active").ToList();

                    for (int n = 0; n < CurrentFileList.Count; n++)
                    {
                        if (!CurrentFile.Contains(CurrentFileList[n].FileName))
                        {
                            CurrentFileList[n].ModifiedBy = User.Identity.Name;
                            CurrentFileList[n].ModifiedDate = DateTime.Now;
                            CurrentFileList[n].Status = "Inactive";
                            _context.Update(CurrentFileList[n]);
                        }
                    }

                    for (int i = 0; i < ProductID.Count; i++)
                    {
                        if (ProductID[i] != null && ProductID[i] != "")
                        {
                            bool IsNew = true;

                            for (int n = 0; n < SalesInvoice.SalesInvoiceDetail.Count; n++)
                            {
                                if (ProductID[i] == SalesInvoice.SalesInvoiceDetail[n].ProductID)
                                {
                                    SalesInvoice.SalesInvoiceDetail[n].UnitPrice = UnitPrice[i];
                                    SalesInvoice.SalesInvoiceDetail[n].Qty = Qty[i];
                                    SalesInvoice.TotalAmount += (UnitPrice[i] * Qty[i]);
                                    _context.Update(SalesInvoice.SalesInvoiceDetail[n]);
                                    IsNew = false;
                                    break;
                                }
                            }

                            if (IsNew)
                            {
                                Product product = _context.Product.Where(e => e.ProductID == ProductID[i]).FirstOrDefault();

                                SalesInvoiceDetail invoiceDetail = new SalesInvoiceDetail();
                                invoiceDetail.InvoiceNo = SalesInvoice.InvoiceNo;
                                invoiceDetail.ProductID = ProductID[i];
                                invoiceDetail.UnitPrice = UnitPrice[i];
                                invoiceDetail.Qty = Qty[i];
                                SalesInvoice.TotalAmount += (UnitPrice[i] * Qty[i]);
                                _context.Add(invoiceDetail);
                            }



                        }
                    }


                    foreach (var file in files)
                    {
                        if (!(file == null || file.Length == 0))
                        {
                            string[] temp = file.FileName.Split(".");
                            string extension = temp[temp.Length - 1];

                            var path = Path.Combine(_context.AppSettings.SingleOrDefaultAsync(a => a.Name == "SalesInvoiceUploadFolder").Result.Value + "\\\\" + invoice.InvoiceNo + "\\\\");

                            if (!System.IO.Directory.Exists(path))
                                System.IO.Directory.CreateDirectory(path);

                            var filename = invoice.InvoiceNo + "_" + (System.IO.Directory.GetFiles(path).Length + 1).ToString() + "." + extension;
                            using (var stream = new FileStream(path + filename, FileMode.Create))
                            {
                                await file.CopyToAsync(stream);
                                SalesInvoiceDocument document = new SalesInvoiceDocument();
                                document.InvoiceNo = invoice.InvoiceNo;
                                document.FileName = invoice.InvoiceNo + "_" + (Directory.GetFiles(path).Length).ToString() + "." + extension;
                                document.Status = "Active";
                                document.CreatedBy = User.Identity.Name;
                                document.CreatedDate = DateTime.Now;
                                _context.Add(document);
                            }
                        }
                    }

                }


                //untrace ticket's SalesInvoice no when cancel
                if (Action == "Cancel")
                {
                    for (int i = 0; i < ProductID.Count; i++)
                    {
                        if (ProductID[i] != null && ProductID[i] != "")
                        {
                            if (InvoiceType[i] == "Ticket")
                            {
                                var ticketDetail = _context.TicketDetail.FirstOrDefault(e => e.TicketID == refID[i] && e.ProductID == ProductID[i]);
                                ticketDetail.SINo = "";
                                ticketDetail.SIFlag = false;
                                _context.Update(ticketDetail);
                            }
                        }
                    }
                }

                _context.Update(SalesInvoice);
                _context.Add(
                    new Log()
                    {
                        EmployeeID = User.Identity.Name,
                        // Entity = User.FindFirstValue(ClaimTypes.System),
                        Controller = ControllerContext.ActionDescriptor.ControllerName,
                        Method = ControllerContext.ActionDescriptor.ActionName,
                        Action = Action,
                        ReferenceID = invoice.InvoiceNo,
                        ActionDate = DateTime.Now
                    }
                );

                await _context.SaveChangesAsync();
                if (Action == "Draft")
                {
                    TempData["Notice"] = "Sales Invoice (" + invoice.InvoiceNo + ") has been saved.";
                    return RedirectToAction(nameof(SalesInvoice_Edit), new { id = dataProtector.Protect(invoice.InvoiceNo) });
                }
                else if (Action == "Submit")
                {
                    TempData["Success"] = "Sales Invoice (" + invoice.InvoiceNo + ") has been created successfully.";
                    return RedirectToAction(nameof(SalesInvoice_Details), new { id = dataProtector.Protect(invoice.InvoiceNo) });
                }
                else if (Action == "Cancel")
                {
                    TempData["Success"] = "Sales Invoice (" + invoice.InvoiceNo + ") has been canceled successfully.";
                    return RedirectToAction(nameof(SalesInvoice_Details), new { id = dataProtector.Protect(invoice.InvoiceNo) });
                }
            }
            else
            {
                var errors = string.Join(" | ", ModelState.Values
                            .SelectMany(v => v.Errors)
                            .Select(e => e.ErrorMessage));
                TempData["Error"] = "Invalid ModelState" + errors;
            }
            //return RedirectToAction(nameof(SalesInvoice_Edit), new { id = dataProtector.Protect(invoice.InvoiceNo) });
            return RedirectToAction(nameof(SalesInvoice_Edit));
        }

        // GET: Sales Invoice/Details/5
        public async Task<IActionResult> SalesInvoice_Details(string id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var invoice = await _context.SalesInvoice
                .Include(p => p.Customer).Include(e => e.SalesInvoiceDetail)
                .SingleOrDefaultAsync(m => m.InvoiceNo == dataProtector.Unprotect(id));

            if (invoice == null)
            {
                return NotFound();
            }
            invoice.ProtectedInvoiceNo = id;
            invoice.SalesInvoiceDocument = _context.SalesInvoiceDocument.Where(e => e.InvoiceNo == invoice.InvoiceNo && e.Status == "Active").ToList();

            foreach (var detail in invoice.SalesInvoiceDetail)
            {
                detail.Product = _context.Product.FirstOrDefault(e => e.ProductID == detail.ProductID);
                detail.ProductUOM = _context.ProductUOM.FirstOrDefault(e => e.OptionID == detail.UOM);
            }

            ViewBag.DocumentURL = _context.AppSettings.SingleOrDefaultAsync(a => a.Name == "SalesInvoiceDisplayFolder").Result.Value;

            return View(invoice);
        }


        // POST: Sales Invoice/SalesInvoice_Details
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> SalesInvoice_Details(string InvoiceNo, string Action, string VoidReason)
        {
            List<string> DOList = new List<String>();

            var payment = await _context.SalesInvoice
                .Include(e => e.SalesInvoiceDetail)
                .SingleOrDefaultAsync(m => m.InvoiceNo == dataProtector.Unprotect(InvoiceNo));

            if (Action == "Void")
            {
                payment.InvoiceStatus = "Void";
                payment.Remark += "Void Reason:" + VoidReason + ". Void by:" + User.Identity.Name + " " + DateTime.Now;
                payment.ModifiedBy = User.Identity.Name;
                payment.ModifiedDate = DateTime.Now;


                foreach (var detail in payment.SalesInvoiceDetail)
                {

                    //untrace Ticket or DeliveryOrder 
                    if (detail.RefID != null && detail.RefID != "")
                    {
                        if (detail.InvoiceType == "Ticket")
                        {
                            var ticketDetail = _context.TicketDetail.FirstOrDefault(e => e.TicketID == detail.RefID && e.ProductID == detail.ProductID);
                            ticketDetail.SINo = "";
                            ticketDetail.SIFlag = false;
                            _context.Update(ticketDetail);
                        }
                        //else if (detail.InvoiceType == "SalesOrder")
                        //{
                        //    var SODetail = _context.SalesOrderDetail.FirstOrDefault(e => e.SalesOrderNo == detail.RefID && e.ProductID == detail.ProductID);
                        //    SODetail.SINo = "";
                        //    SODetail.SIFlag = false;
                        //    _context.Update(SODetail);
                        //}
                        else if (detail.InvoiceType == "DO")
                        {
                            var DODetail = _context.DeliveryOrderDetail.FirstOrDefault(e => e.DeliveryOrderNo == detail.RefID && e.ProductID == detail.ProductID);
                            DODetail.SINo = "";
                            DODetail.SIFlag = false;
                            _context.Update(DODetail);

                            DOList.Add(detail.RefID);
                        }
                    }

                }
            }

            _context.Update(payment);

            _context.Add(
                new Log()
                {
                    EmployeeID = User.Identity.Name,
                    SiteID = User.FindFirstValue("SiteID"),
                    Controller = ControllerContext.ActionDescriptor.ControllerName,
                    Method = ControllerContext.ActionDescriptor.ActionName,
                    Action = Action,
                    ReferenceID = payment.InvoiceNo,
                    ActionDate = DateTime.Now
                }
            );

            await _context.SaveChangesAsync();

            if (Action == "Void")
            {
                //DO status change back to Confirmed after SalesInvoice Untrace
                List<string> DODistinct = new List<String>();
                DODistinct.AddRange(DOList.Distinct());

                for (int i = 0; i < DODistinct.Count; i++)
                {
                    if (DODistinct[i] != null && DODistinct[i] != "")
                    {
                        var Result = _context.DeliveryOrderDetail.Any(e => e.SIFlag == false && e.DeliveryOrderNo == DODistinct[i]);
                        if (Result)
                        {
                            var DO = _context.DeliveryOrder.FirstOrDefault(e => e.DeliveryOrderNo == DODistinct[i]);
                            DO.DeliveryOrderStatus = "Confirmed";
                            _context.Update(DO);
                        }
                    }

                }
                await _context.SaveChangesAsync();
            }

            if (Action == "Void")
            {
                TempData["Success"] = "Sales Invoice (" + payment.InvoiceNo + ") has been void successfully.";
                return RedirectToAction(nameof(SalesInvoice_Details), new { id = dataProtector.Protect(payment.InvoiceNo) });
            }

            return RedirectToAction(nameof(SalesInvoice_Details));

        }


        // GET: PrintInvoice
        public ActionResult PrintInvoice(string id, string ReportType)
        {
            if (id == null)
            {
                return NotFound();
            }
            var key = _context.AppSettings.SingleOrDefault(a => a.Name == "Key");
            ViewBag.ID = Common.EncryptString(dataProtector.Unprotect(id), key.Value);
            ViewBag.ReportType = ReportType;

            var ReportServer = _context.AppSettings.SingleOrDefault(a => a.Name == "ReportServer");
            ViewBag.ReportServer = ReportServer.Value;

            return View();
        }


        // GET: Customer Payments
        public ActionResult CustomerPayment_Index()
        {
            PrepareCustomerPaymentIndexView();
            return View(new List<InPayment>());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> CustomerPayment_Index(string CustomerID, string PaymentNo, string Mode)
        {
            var payments = await _context.InPayment
                        .Where(e => e.CustomerID == (CustomerID != null ? CustomerID : e.CustomerID)
                        && e.PaymentNo.Contains((PaymentNo != null ? PaymentNo : e.PaymentNo))
                        && e.PaymentMode == (Mode != null ? Mode : e.PaymentMode))
                        .ToListAsync();

            foreach (var pay in payments)
            {
                pay.Customer = _context.Customer.FirstOrDefault(e => e.CustomerID == pay.CustomerID);
                pay.ProtectedPaymentNo = dataProtector.Protect(pay.PaymentNo);
            }
            PrepareCustomerPaymentIndexView();
            return View(payments);
        }

        private void PrepareCustomerPaymentIndexView()
        {
            var claimSite = User.FindFirstValue("SiteID");
            List<SelectListItem> MenuList = new List<SelectListItem>();
            var psSQL = from s in _context.Customer
                        where s.SiteID == claimSite || s.SiteID == null
                        orderby s.CustomerCode ascending
                        select s;
            var psList = psSQL.ToList();
            if (psList.Count > 0)
            {
                for (int i = 0; i < psList.Count; i++)
                {
                    SelectListItem item = new SelectListItem { Value = psList[i].CustomerID, Text = psList[i].CustomerCode + " - " + psList[i].CustomerName };
                    MenuList.Add(item);
                }
            }
            ViewBag.CustomerList = MenuList;
        }



        // GET: Customer Payments/Create
        public IActionResult CustomerPayment_Create()
        {
            PrepareCustomerPaymentIndexView();
            return View();
        }

        // POST: Customer Payments/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> CustomerPayment_Create([Bind("CustomerID,PaymentTotal,PaymentMode,ReferenceNo,Remark")] InPayment payment, List<string> InvoiceNo, List<decimal> PaymentAmt, List<IFormFile> files, string PaymentDate)
        {
            DateTime dtNow = DateTime.Now;
            if (ModelState.IsValid)
            {
                var c = _context.Customer.Where(x => x.CustomerID == payment.CustomerID).FirstOrDefault();
                if (c == null)
                {
                    TempData["Error"] = "Customer not found.";
                    return RedirectToAction(nameof(CustomerPayment_Index));
                }

                bool Proceed = true;

                if (payment.PaymentMode == "Cheque" && _context.InPayment.Any(e => e.ReferenceNo == payment.ReferenceNo && e.PaymentStatus == "Paid"))
                    Proceed = false;

                if (Proceed)
                {
                    IDSetupController id = new IDSetupController(_context);
                    payment.SiteID = User.FindFirstValue("SiteID");
                    payment.PaymentNo = id.InternalGetNewID("InPayment");
                    payment.PaymentDate = DateTime.ParseExact(PaymentDate, "dd/MM/yyyy", null);
                    payment.CreatedBy = User.Identity.Name;
                    payment.CreatedDate = DateTime.Now;
                    payment.PaymentStatus = "Paid";
                    if (payment.PaymentMode != "Cheque") 
                        payment.ReferenceNo = null;
                    _context.Add(payment);

                    if (payment.PaymentMode == "Credit")
                    {
                        var dPaymentTotal = Math.Round(Convert.ToDouble(payment.PaymentTotal), 2, MidpointRounding.AwayFromZero);

                        if (dPaymentTotal > c.Credit)
                        {
                            TempData["Error"] = "Insufficient credit.";
                            return RedirectToAction(nameof(CustomerPayment_Index));
                        }

                        //deduct credit and insert Tan
                        var dBefore = c.Credit;
                        c.Credit = Math.Round(c.Credit - dPaymentTotal, 2, MidpointRounding.AwayFromZero);
                        var dAfter = c.Credit;

                        var t = new CreditTran
                        {
                            Id = Guid.NewGuid().ToString(),
                            CustomerId = c.CustomerID,
                            Datetime = dtNow,
                            Action = "SpendCredit",
                            Remark = null,
                            Credit = dPaymentTotal * -1,
                            Before = dBefore,
                            After = dAfter
                        };
                        _context.CreditTrans.Add(t);
                    }

                    for (int i = 0; i < InvoiceNo.Count; i++)
                    {
                        if (InvoiceNo[i] != null && InvoiceNo[i] != "" && PaymentAmt[i] > 0)
                        {
                            InPaymentDetail paymentDetail = new InPaymentDetail();
                            paymentDetail.RefNo = InvoiceNo[i];
                            paymentDetail.PaymentAmount = PaymentAmt[i];
                            paymentDetail.PaymentNo = payment.PaymentNo;

                            var SalesInvoice = _context.SalesInvoice.FirstOrDefault(e => e.InvoiceNo == InvoiceNo[i] && e.InvoiceStatus == "Payment");// && e.InvoiceStage == "Payment");
                            SalesInvoice.PaidAmount += PaymentAmt[i];
                            SalesInvoice.ModifiedBy = User.Identity.Name;
                            SalesInvoice.ModifiedDate = DateTime.Now;

                            if (SalesInvoice.TotalAmount - SalesInvoice.PaidAmount == 0)
                                SalesInvoice.InvoiceStatus = "Paid";

                            _context.Add(paymentDetail);
                            _context.Update(SalesInvoice);
                        }
                    }

                    foreach (var file in files)
                    {
                        if (!(file == null || file.Length == 0))
                        {
                            string[] temp = file.FileName.Split(".");
                            string extension = temp[temp.Length - 1];

                            var path = Path.Combine(_context.AppSettings.SingleOrDefaultAsync(a => a.Name == "InPaymentUploadFolder").Result.Value + "\\\\" + payment.PaymentNo + "\\\\");
                            string directoryName = Path.GetDirectoryName(path);

                            if (!System.IO.Directory.Exists(path))
                                System.IO.Directory.CreateDirectory(path);

                            path += payment.PaymentNo + "_" + (Directory.GetFiles(directoryName).Length + 1).ToString() + "." + extension;

                            using (var stream = new FileStream(path, FileMode.Create))
                            {
                                await file.CopyToAsync(stream);
                                InPaymentDocument document = new InPaymentDocument();
                                document.PaymentNo = payment.PaymentNo;
                                document.FileName = payment.PaymentNo + "_" + (Directory.GetFiles(directoryName).Length).ToString() + "." + extension;
                                document.Status = "Active";
                                document.CreatedBy = User.Identity.Name;
                                document.CreatedDate = DateTime.Now;
                                _context.Add(document);
                            }
                        }
                    }

                    if (Proceed)
                    {
                        _context.Add(
                            new Log()
                            {
                                EmployeeID = User.Identity.Name,
                                SiteID = User.FindFirstValue("SiteID"),
                                Controller = ControllerContext.ActionDescriptor.ControllerName,
                                Method = ControllerContext.ActionDescriptor.ActionName,
                                Action = "New Customer Payment",
                                ReferenceID = payment.PaymentNo,
                                ActionDate = DateTime.Now
                            }
                        );
                        await _context.SaveChangesAsync();
                        TempData["Success"] = "Payment (" + payment.PaymentNo + ") has been created successfully.";
                        return RedirectToAction(nameof(CustomerPayment_Details), new { id = dataProtector.Protect(payment.PaymentNo) });
                    }
                    else
                    {
                        TempData["Error"] = "Unable to create payment. There are changes in invoice payable amount.";
                        return RedirectToAction(nameof(CustomerPayment_Index));
                    }
                }
                else
                {
                    TempData["Notice"] = "Unable to create payment. Cheque No is duplicated.";
                    return RedirectToAction(nameof(CustomerPayment_Index));
                }
            }

            return View(payment);
        }

        // GET: Customer Payment/Details/5
        public async Task<IActionResult> CustomerPayment_Details(string id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var payment = await _context.InPayment
                .Include(p => p.Customer).Include(e => e.InPaymentDetail)
                .SingleOrDefaultAsync(m => m.PaymentNo == dataProtector.Unprotect(id));

            if (payment == null)
            {
                return NotFound();
            }

            payment.ProtectedPaymentNo = id;
            payment.InPaymentDocument = _context.InPaymentDocument.Where(e => e.PaymentNo == payment.PaymentNo && e.Status == "Active").ToList();

            foreach (var detail in payment.InPaymentDetail)
            {
                detail.SalesInvoice = _context.SalesInvoice.FirstOrDefault(e => e.InvoiceNo == detail.RefNo);
                detail.PaymentVoucher = _context.PaymentVoucher.FirstOrDefault(e => e.PVNo == detail.RefNo);  // cater for Advance receipt
            }

            ViewBag.DocumentURL = _context.AppSettings.SingleOrDefaultAsync(a => a.Name == "InPaymentDisplayFolder").Result.Value;
            return View(payment);
        }

        //private bool PaymentExists(string id)
        //{
        //    return _context.InPayment.Any(e => e.PaymentNo == id);
        //}

        public FileContentResult CustomerReceipt_Pdf(string id)
        {
            id = dataProtector.Unprotect(id);

            Byte[] bytes;
            using (var ms = new MemoryStream())
            //using (var doc = new Document(pagesize == "A5" ? PageSize.A5.Rotate() : PageSize.A4))
            using (var doc = new Document(PageSize.A4))
            using (var writer = PdfWriter.GetInstance(doc, ms))
            {
                try
                {
                    var site = _context.Site.FirstOrDefault();
                    var pmt = _context.InPayment
                        .Include(x => x.Customer)
                        .Where(x => x.PaymentNo == id).FirstOrDefault();
                    var details = _context.InPaymentDetail.Where(x => x.PaymentNo == pmt.PaymentNo).ToList();

                    if (site == null || pmt == null)
                        return null;

                    CustomerReceipt_Pdf_Helper helper = new CustomerReceipt_Pdf_Helper();
                    writer.PageEvent = helper;

                    doc.Open();
                    PdfPTable table, childtable, ctable;
                    PdfPCell cell, titlecell, datacell, datacell2, ccell;
                    Paragraph space = new Paragraph("\n", new Font(Font.FontFamily.HELVETICA, 12f, Font.NORMAL));
                    PdfPCell lineCell = new PdfPCell(new Phrase("")); lineCell.BorderColorBottom = new BaseColor(System.Drawing.Color.Black); lineCell.BorderWidthBottom = 0.5f;
                    PdfPTable line = new PdfPTable(1); line.WidthPercentage = 100; line.AddCell(lineCell);
                    PdfPCell lineCellGrey = new PdfPCell(new Phrase("")); lineCellGrey.BorderColor = new BaseColor(System.Drawing.Color.LightGray); lineCellGrey.BorderWidthBottom = 1f;
                    PdfPTable lineGrey = new PdfPTable(1); lineGrey.WidthPercentage = 100; lineGrey.AddCell(lineCellGrey);
                    Font font20b = new Font(Font.FontFamily.HELVETICA, 20f, Font.BOLD);
                    Font font15b = new Font(Font.FontFamily.HELVETICA, 15f, Font.BOLD);
                    Font font12b = new Font(Font.FontFamily.HELVETICA, 12f, Font.BOLD);
                    Font font10 = new Font(Font.FontFamily.HELVETICA, 10f, Font.NORMAL);
                    Font font10b = new Font(Font.FontFamily.HELVETICA, 10f, Font.BOLD);
                    Font font10grey = new Font(Font.FontFamily.HELVETICA, 10f, Font.NORMAL, new BaseColor(System.Drawing.Color.Gray));
                    Font font8 = new Font(Font.FontFamily.HELVETICA, 8f, Font.NORMAL);
                    Font font8b = new Font(Font.FontFamily.HELVETICA, 8f, Font.BOLD);
                    Font font6 = new Font(Font.FontFamily.HELVETICA, 6f, Font.NORMAL);
                    Font font6b = new Font(Font.FontFamily.HELVETICA, 6f, Font.BOLD);
                    BaseColor colorBlack = new BaseColor(System.Drawing.Color.Black);

                    int iNoOfCopy = 1;
                    int iPrinted = 0;
                    var setting = _context.AppSettings.Where(x => x.Name == "CustomerPaymentReceipt_NoOfCopy").FirstOrDefault();
                    if (setting != null) iNoOfCopy = Convert.ToInt32(setting.Value);
                        
                    Print:

                    //r1
                    table = new PdfPTable(2); table.WidthPercentage = 100;
                    cell = new PdfPCell(); cell.HorizontalAlignment = 0; cell.BorderColor = colorBlack; cell.Border = PdfPCell.BOTTOM_BORDER;
                    ctable = new PdfPTable(1); ctable.WidthPercentage = 100;
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(site.SiteName, font12b); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(site.RegNo, font8); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(site.Address1, font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(site.Address2, font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(site.PhoneNo, font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    cell.AddElement(ctable);
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Payment Receipt", font20b));
                    cell.HorizontalAlignment = 2; cell.BorderColor = colorBlack; cell.Border = PdfPCell.BOTTOM_BORDER;
                    table.AddCell(cell);
                    doc.Add(table);
                    doc.Add(space);

                    //r2
                    table = new PdfPTable(3); table.WidthPercentage = 100;
                    cell = new PdfPCell(); cell.Colspan = 2; cell.HorizontalAlignment = 0; cell.BorderColor = colorBlack; cell.Border = 0;
                    cell.AddElement(new Phrase("Received from :", font10));
                    cell.AddElement(new Phrase(pmt.Customer.CustomerName, font10b));
                    table.AddCell(cell);
                    cell = new PdfPCell(); cell.Colspan = 1; cell.BorderColor = colorBlack; cell.Border = 0;
                    ctable = new PdfPTable(2); ctable.HorizontalAlignment = 2; ctable.WidthPercentage = 100;
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase("PAYMENT NO :", font10b); ccell.HorizontalAlignment = 2; ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(pmt.PaymentNo, font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase("DATE :", font10); ccell.HorizontalAlignment = 2; ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(pmt.PaymentDate.ToString("dd-MMM-yyyy"), font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase("MODE :", font10); ccell.HorizontalAlignment = 2; ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(pmt.PaymentMode, font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    cell.AddElement(ctable);
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("TEL: FAX: ", font10)); cell.Colspan = 3; cell.BorderColor = colorBlack; cell.Border = 0;
                    table.AddCell(cell);
                    doc.Add(table);
                    doc.Add(space);

                    //r3
                    table = new PdfPTable(25); table.WidthPercentage = 100;
                    cell = new PdfPCell(new Phrase("#", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 3; cell.HorizontalAlignment = 0; table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("INVOICE NO", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 8; cell.HorizontalAlignment = 0; table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("INVOICE DATE", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 6; cell.HorizontalAlignment = 1; table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("GRAND TOTAL", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 4; cell.HorizontalAlignment = 2; table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("PAID", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 4; cell.HorizontalAlignment = 2; table.AddCell(cell);
                    
                    for (int i = 0; i < details.Count; i++)
                    {
                        var inv = _context.SalesInvoice.Where(x => x.InvoiceNo == details[i].RefNo).FirstOrDefault();
                        if (inv != null)
                        {
                            var iRepeat = 1;
                            for (int j = 0; j < iRepeat; j++)
                            {
                                cell = new PdfPCell(new Phrase((i + 1).ToString(), font10));
                                cell.Colspan = 3; cell.HorizontalAlignment = 0; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                                cell = new PdfPCell(new Phrase(details[i].RefNo, font10));
                                cell.Colspan = 8; cell.HorizontalAlignment = 0; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                                cell = new PdfPCell(new Phrase(inv.InvoiceDate.ToString("dd-MMM-yyyy"), font10));
                                cell.Colspan = 6; cell.HorizontalAlignment = 1; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                                cell = new PdfPCell(new Phrase(inv.TotalAmount.ToString("N"), font10));
                                cell.Colspan = 4; cell.HorizontalAlignment = 2; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                                cell = new PdfPCell(new Phrase(details[i].PaymentAmount.ToString("N"), font10));
                                cell.Colspan = 4; cell.HorizontalAlignment = 2; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                            }
                        }
                    }
                    doc.Add(table);
                    doc.Add(space);
                    doc.Add(line);
                    doc.Add(space);

                    //r4
                    table = new PdfPTable(8); table.WidthPercentage = 100;
                    cell = new PdfPCell(); cell.Colspan = 3; cell.HorizontalAlignment = 0; cell.BorderColor = colorBlack; cell.Border = 0;
                    Paragraph p = new Paragraph("Received By:                  ");
                    LineSeparator ls = new LineSeparator(1, 100, null, Element.ALIGN_BOTTOM, -10);
                    p.Add(ls);
                    cell.AddElement(p);
                    table.AddCell(cell);
                    cell = new PdfPCell(); cell.Colspan = 2; cell.HorizontalAlignment = 0; cell.BorderColor = colorBlack; cell.Padding = 10; cell.Border = 0;
                    table.AddCell(cell);
                    cell = new PdfPCell(); cell.Colspan = 3; cell.HorizontalAlignment = 0; cell.BorderColor = colorBlack; cell.Padding = 5; /*cell.Border = 0;*/
                    ctable = new PdfPTable(2); ctable.HorizontalAlignment = 2;
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase("PAID AMOUNT", font10); ccell.HorizontalAlignment = 0; ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase("RM " + pmt.PaymentTotal.ToString("N"), font12b); ccell.HorizontalAlignment = 2; ccell.Border = 0;
                    ctable.AddCell(ccell);
                    cell.AddElement(ctable);
                    table.AddCell(cell);
                    doc.Add(table);

                    iPrinted++;

                    if (iPrinted < iNoOfCopy)
                    {
                        if (details.Count <= 5)
                        {
                            doc.Add(space);
                            doc.Add(line);
                            doc.Add(space);
                        }
                        else
                        {
                            doc.NewPage();
                        }

                        goto Print;
                    }   
                }
                finally
                {
                    doc.Close();
                }
                bytes = ms.ToArray();
            }

            return new FileContentResult(bytes, "application/pdf");
        }

        class CustomerReceipt_Pdf_Helper : PdfPageEventHelper
        {
            private PdfContentByte cb, cb2;
            private List<PdfTemplate> templates, templates2;
            public CustomerReceipt_Pdf_Helper()
            {
                templates = new List<PdfTemplate>();
                templates2 = new List<PdfTemplate>();
            }

            public override void OnEndPage(PdfWriter writer, Document document)
            {
                base.OnEndPage(writer, document);

                ////draw red border for troubleshoot layout
                //var content = writer.DirectContent;
                //var pageBorderRect = new Rectangle(document.PageSize);
                //pageBorderRect.Left += document.LeftMargin;
                //pageBorderRect.Right -= document.RightMargin;
                //pageBorderRect.Top -= document.TopMargin;
                //pageBorderRect.Bottom += document.BottomMargin;
                //content.SetColorStroke(BaseColor.RED);
                //content.Rectangle(pageBorderRect.Left, pageBorderRect.Bottom, pageBorderRect.Width, pageBorderRect.Height);
                //content.Stroke();

                //page number
                cb = writer.DirectContentUnder;
                PdfTemplate templateM = cb.CreateTemplate(50, 50);
                templates.Add(templateM);
                //string pageText = "Page " + writer.PageNumber + " of ";
                string pageText = "Page " + writer.PageNumber;
                BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                float len = bf.GetWidthPoint(pageText, 10);
                cb.BeginText();
                cb.SetFontAndSize(bf, 10);
                cb.SetTextMatrix(document.LeftMargin, document.PageSize.GetBottom(document.BottomMargin) - 10);
                cb.ShowText(pageText);
                cb.EndText();
                //cb.AddTemplate(templateM, document.LeftMargin + len, document.PageSize.GetBottom(document.BottomMargin) - 10);

                cb2 = writer.DirectContentUnder;
                PdfTemplate templateM2 = cb2.CreateTemplate(50, 50);
                templates2.Add(templateM2);
                string pageText2 = DateTime.Now.ToString("yyyy-MM-dd");
                BaseFont bf2 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                float len2 = bf2.GetWidthPoint(pageText2, 10);
                cb2.BeginText();
                cb2.SetFontAndSize(bf2, 10);
                cb2.SetTextMatrix(document.PageSize.GetRight(document.RightMargin) - 50, document.PageSize.GetBottom(document.BottomMargin) - 10);
                cb2.ShowText(pageText2);
                cb2.EndText();
                //cb2.AddTemplate(templateM2, document.RightMargin - len2, document.PageSize.GetBottom(document.BottomMargin) - 10);
            }

            public override void OnCloseDocument(PdfWriter writer, Document document)
            {
                base.OnCloseDocument(writer, document);
                //BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                //foreach (PdfTemplate item in templates)
                //{
                //    item.BeginText();
                //    item.SetFontAndSize(bf, 10);
                //    item.SetTextMatrix(0, 0);
                //    item.ShowText("" + writer.PageNumber);
                //    item.EndText();
                //}
            }
        }

        public FileContentResult SalesInvoice_Pdf(string id)
        {
            id = dataProtector.Unprotect(id);

            Byte[] bytes;
            using (var ms = new MemoryStream())
            //using (var doc = new Document(pagesize == "A5" ? PageSize.A5.Rotate() : PageSize.A4))
            using (var doc = new Document(PageSize.A4))
            using (var writer = PdfWriter.GetInstance(doc, ms))
            {
                try
                {
                    var site = _context.Site.FirstOrDefault();
                    var pmt = _context.SalesInvoice
                        .Include(x => x.Customer)
                        .Where(x => x.InvoiceNo == id).FirstOrDefault();
                    var details = _context.SalesInvoiceDetail.Where(x => x.InvoiceNo == pmt.InvoiceNo).ToList();

                    if (site == null || pmt == null)
                        return null;

                    SalesInvoice_Pdf_Helper helper = new SalesInvoice_Pdf_Helper();
                    writer.PageEvent = helper;

                    doc.Open();
                    PdfPTable table, childtable, ctable, boxTable;
                    PdfPCell cell, titlecell, datacell, datacell2, ccell, boxCell;
                    Paragraph space = new Paragraph("\n", new Font(Font.FontFamily.TIMES_ROMAN, 12f, Font.NORMAL));
                    PdfPCell lineCell = new PdfPCell(new Phrase("")); lineCell.BorderColorBottom = new BaseColor(System.Drawing.Color.Black); lineCell.BorderWidthBottom = 0.5f;
                    PdfPTable line = new PdfPTable(1); line.WidthPercentage = 100; line.AddCell(lineCell);
                    PdfPCell lineCellGrey = new PdfPCell(new Phrase("")); lineCellGrey.BorderColor = new BaseColor(System.Drawing.Color.LightGray); lineCellGrey.BorderWidthBottom = 1f;
                    PdfPTable lineGrey = new PdfPTable(1); lineGrey.WidthPercentage = 100; lineGrey.AddCell(lineCellGrey);
                    Font font20b = new Font(Font.FontFamily.TIMES_ROMAN, 20f, Font.BOLD);
                    Font font15b = new Font(Font.FontFamily.TIMES_ROMAN, 15f, Font.BOLD);
                    Font font12b = new Font(Font.FontFamily.TIMES_ROMAN, 12f, Font.BOLD);
                    Font font10 = new Font(Font.FontFamily.TIMES_ROMAN, 10f, Font.NORMAL);
                    Font font10b = new Font(Font.FontFamily.TIMES_ROMAN, 10f, Font.BOLD);
                    Font font10grey = new Font(Font.FontFamily.TIMES_ROMAN, 10f, Font.NORMAL, new BaseColor(System.Drawing.Color.Gray));
                    Font font8 = new Font(Font.FontFamily.TIMES_ROMAN, 8f, Font.NORMAL);
                    Font font8b = new Font(Font.FontFamily.TIMES_ROMAN, 8f, Font.BOLD);
                    Font font6 = new Font(Font.FontFamily.TIMES_ROMAN, 6f, Font.NORMAL);
                    Font font6b = new Font(Font.FontFamily.TIMES_ROMAN, 6f, Font.BOLD);
                    BaseColor colorBlack = new BaseColor(System.Drawing.Color.Black);

                    int iNoOfCopy = 1;
                    int iPrinted = 0;

                Print:

                    //r1
                    table = new PdfPTable(2); table.WidthPercentage = 100;
                    cell = new PdfPCell(); cell.HorizontalAlignment = 0; cell.BorderColor = colorBlack; cell.Border = PdfPCell.BOTTOM_BORDER;
                    ctable = new PdfPTable(1); ctable.WidthPercentage = 100;
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(site.SiteName, font12b); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(site.RegNo, font8); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(site.Address1, font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(site.Address2, font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(site.PhoneNo, font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    cell.AddElement(ctable);
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("SALES INVOICE", font20b));
                    cell.HorizontalAlignment = 2; cell.BorderColor = colorBlack; cell.Border = PdfPCell.BOTTOM_BORDER;
                    table.AddCell(cell);
                    doc.Add(table);

                    //r2
                    var ShowCustCode = _context.AppSettings.SingleOrDefault(a => a.Name == "ShowCustCode_SalesInvoice");
                    ViewBag.ShowCustCode = ShowCustCode.Value;

                    table = new PdfPTable(3); table.WidthPercentage = 100;
                    cell = new PdfPCell(); cell.Colspan = 2; cell.HorizontalAlignment = 0; cell.BorderColor = colorBlack; cell.Border = 0;
                    cell.AddElement(new Phrase("BILL TO :", font10));
                    if (ViewBag.ShowCustCode == "1")
                    {
                        cell.AddElement(new Phrase(pmt.Customer.CustomerCode + " " + pmt.Customer.CustomerName, font10b));
                    }
                    else
                    {
                        cell.AddElement(new Phrase(pmt.Customer.CustomerName, font10b));
                    }
                    table.AddCell(cell);
                    cell = new PdfPCell(); cell.Colspan = 1; cell.BorderColor = colorBlack; cell.Border = 0;
                    ctable = new PdfPTable(2); ctable.HorizontalAlignment = 2; ctable.WidthPercentage = 100;
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase("INV NO :", font12b); ccell.HorizontalAlignment = 2; ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(pmt.InvoiceNo, font12b); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase("DATE :", font10); ccell.HorizontalAlignment = 2; ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(pmt.InvoiceDate.ToString("dd-MMM-yyyy"), font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase("P.O. NO :", font10); ccell.HorizontalAlignment = 2; ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase(pmt.ReferenceNo, font10); ccell.Border = 0;
                    ctable.AddCell(ccell);
                    cell.AddElement(ctable);
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(pmt.Customer.BillingAddress1, font10)); cell.Colspan = 3; cell.BorderColor = colorBlack; cell.Border = 0;
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(pmt.Customer.BillingAddress2, font10)); cell.Colspan = 3; cell.BorderColor = colorBlack; cell.Border = 0;
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(pmt.Customer.BillingPostcode + " " + pmt.Customer.BillingCity, font10)); cell.Colspan = 3; cell.BorderColor = colorBlack; cell.Border = 0;
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(pmt.Customer.BillingState + ", " + pmt.Customer.BillingCountry, font10)); cell.Colspan = 3; cell.BorderColor = colorBlack; cell.Border = 0;
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("TEL: " + pmt.Customer.BillingPhone + "   " + "FAX: " + pmt.Customer.BillingFax, font10)); cell.Colspan = 3; cell.BorderColor = colorBlack; cell.Border = 0;
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Attn: " + pmt.Customer.BillingPIC, font10)); cell.Colspan = 3; cell.BorderColor = colorBlack; cell.Border = 0;
                    table.AddCell(cell);
                    doc.Add(table);
                    doc.Add(space);

                    //r3
                    table = new PdfPTable(25); table.WidthPercentage = 100;
                    cell = new PdfPCell(new Phrase("#", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 1; cell.HorizontalAlignment = 0; table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("PRODUCT", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 4; cell.HorizontalAlignment = 0; table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("DESCRIPTION", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 9; cell.HorizontalAlignment = 0; table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("QTY", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 3; cell.HorizontalAlignment = 0; table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("UOM", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 2; cell.HorizontalAlignment = 0; table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("U.PRICE", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 3; cell.HorizontalAlignment = 0; table.AddCell(cell);
                    cell = new PdfPCell(new Phrase("NET AMT", font10b)); cell.Border = PdfPCell.TOP_BORDER | PdfPCell.BOTTOM_BORDER;
                    cell.Colspan = 3; cell.HorizontalAlignment = 0; table.AddCell(cell);

                    for (int i = 0; i < details.Count; i++)
                    {
                        var inv = _context.SalesInvoice.Where(x => x.InvoiceNo == details[i].InvoiceNo).FirstOrDefault();
                        var prod = _context.Product.Where(x => x.ProductID == details[i].ProductID).FirstOrDefault();
                        var uom = _context.ProductUOM.Where(x => x.OptionID == details[i].UOM).FirstOrDefault();
                        if (inv != null)
                        {
                            var iRepeat = 1;
                            for (int j = 0; j < iRepeat; j++)
                            {
                                cell = new PdfPCell(new Phrase((i + 1).ToString(), font10));
                                cell.Colspan = 1; cell.HorizontalAlignment = 0; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                                cell = new PdfPCell(new Phrase(prod.ProductCode, font10));
                                cell.Colspan = 4; cell.HorizontalAlignment = 0; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                                cell = new PdfPCell(new Phrase(details[i].Description, font10));
                                cell.Colspan = 9; cell.HorizontalAlignment = 0; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                                cell = new PdfPCell(new Phrase(details[i].Qty.ToString("0.####"), font10));
                                cell.Colspan = 3; cell.HorizontalAlignment = 0; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                                cell = new PdfPCell(new Phrase(uom.OptionName, font10));
                                cell.Colspan = 2; cell.HorizontalAlignment = 0; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                                cell = new PdfPCell(new Phrase(details[i].UnitPrice.ToString("N3"), font10));
                                cell.Colspan = 3; cell.HorizontalAlignment = 0; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                                cell = new PdfPCell(new Phrase(details[i].SubTotal.ToString("N"), font10));
                                cell.Colspan = 3; cell.HorizontalAlignment = 0; cell.Border = 0; cell.Padding = 5; table.AddCell(cell);
                            }
                        }
                    }
                    doc.Add(table);
                    doc.Add(space);
                    doc.Add(line);

                    //r4
                    table = new PdfPTable(8); table.WidthPercentage = 100;
                    cell = new PdfPCell(); cell.Colspan = 4; cell.HorizontalAlignment = 0; cell.BorderColor = colorBlack; cell.Border = 0;
                    cell.AddElement(new Phrase("Notes: 1. All cheques should be crossed and made payable to", font10));
                    cell.AddElement(new Phrase(site.SiteName, font10));
                    cell.AddElement(new Phrase(site.BankAccount, font10));
                    cell.AddElement(new Phrase("2. Goods sold are neither returnable nor refundable. Otherwise a cancellation fee of 20 % on purchase price will be imposed.", font10));
                    //doc.Add(space);
                    doc.Add(space);
                    cell.AddElement(new Phrase("                               ", font10));
                    cell.AddElement(new Phrase("                               ", font10));
                    cell.AddElement(new Phrase("                               ", font10));
                    cell.AddElement(new Phrase("___________________________________", font10));
                    cell.AddElement(new Phrase("              Authorised Signature", font12b));
                    //cell.AddElement(new Phrase(pmt.Customer.CustomerName, font10b));
                    //LineSeparator ls = new LineSeparator(1, 100, null, Element.ALIGN_BOTTOM, -10);
                    //Paragraph p = new Paragraph("Authorised Signature:                  ");
                    //p.Add(ls);
                    //cell.AddElement(p);
                    table.AddCell(cell);
                    cell = new PdfPCell(); cell.Colspan = 1; cell.HorizontalAlignment = 0; cell.BorderColor = colorBlack; cell.Padding = 10; cell.Border = 0;
                    table.AddCell(cell);
                    cell = new PdfPCell(); cell.Colspan = 3; cell.HorizontalAlignment = 0; cell.BorderColor = colorBlack; cell.Padding = 5; cell.Border = 0;

                    boxTable = new PdfPTable(2); boxTable.WidthPercentage = 100;
                    boxCell = new PdfPCell(); boxCell.Colspan = 2; boxCell.HorizontalAlignment = 0; boxCell.BorderColor = colorBlack; boxCell.Padding = 5;
                    
                    ctable = new PdfPTable(2); ctable.HorizontalAlignment = 0; ctable.WidthPercentage = 100;
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase("NET AMOUNT", font10); ccell.HorizontalAlignment = 0; ccell.Border = 0;
                    ctable.AddCell(ccell);
                    ccell = new PdfPCell(); ccell.Phrase = new Phrase("RM " + pmt.TotalAmount.ToString("N"), font12b); ccell.HorizontalAlignment = 2; ccell.Border = 0;
                    ctable.AddCell(ccell);

                    boxCell.AddElement(ctable);
                    boxTable.AddCell(boxCell);
                    cell.AddElement(boxTable);

                    table.AddCell(cell);
                    doc.Add(table);

                    iPrinted++;

                    if (iPrinted < iNoOfCopy)
                    {
                        if (details.Count <= 5)
                        {
                            doc.Add(space);
                            doc.Add(line);
                            doc.Add(space);
                        }
                        else
                        {
                            doc.NewPage();
                        }

                        goto Print;
                    }
                }
                finally
                {
                    doc.Close();
                }
                bytes = ms.ToArray();
            }

            return new FileContentResult(bytes, "application/pdf");
        }

        class SalesInvoice_Pdf_Helper : PdfPageEventHelper
        {
            private PdfContentByte cb, cb2;
            private List<PdfTemplate> templates, templates2;
            public SalesInvoice_Pdf_Helper()
            {
                templates = new List<PdfTemplate>();
                templates2 = new List<PdfTemplate>();
            }

            public override void OnEndPage(PdfWriter writer, Document document)
            {
                base.OnEndPage(writer, document);

                ////draw red border for troubleshoot layout
                //var content = writer.DirectContent;
                //var pageBorderRect = new Rectangle(document.PageSize);
                //pageBorderRect.Left += document.LeftMargin;
                //pageBorderRect.Right -= document.RightMargin;
                //pageBorderRect.Top -= document.TopMargin;
                //pageBorderRect.Bottom += document.BottomMargin;
                //content.SetColorStroke(BaseColor.RED);
                //content.Rectangle(pageBorderRect.Left, pageBorderRect.Bottom, pageBorderRect.Width, pageBorderRect.Height);
                //content.Stroke();

                //page number
                cb = writer.DirectContentUnder;
                PdfTemplate templateM = cb.CreateTemplate(50, 50);
                templates.Add(templateM);
                //string pageText = "Page " + writer.PageNumber + " of ";
                string pageText = "Page " + writer.PageNumber;
                BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                float len = bf.GetWidthPoint(pageText, 10);
                cb.BeginText();
                cb.SetFontAndSize(bf, 10);
                cb.SetTextMatrix(document.LeftMargin, document.PageSize.GetBottom(document.BottomMargin) - 10);
                cb.ShowText(pageText);
                cb.EndText();
                //cb.AddTemplate(templateM, document.LeftMargin + len, document.PageSize.GetBottom(document.BottomMargin) - 10);

                cb2 = writer.DirectContentUnder;
                PdfTemplate templateM2 = cb2.CreateTemplate(50, 50);
                templates2.Add(templateM2);
                string pageText2 = DateTime.Now.ToString("yyyy-MM-dd");
                BaseFont bf2 = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                float len2 = bf2.GetWidthPoint(pageText2, 10);
                cb2.BeginText();
                cb2.SetFontAndSize(bf2, 10);
                cb2.SetTextMatrix(document.PageSize.GetRight(document.RightMargin) - 50, document.PageSize.GetBottom(document.BottomMargin) - 10);
                cb2.ShowText(pageText2);
                cb2.EndText();
                //cb2.AddTemplate(templateM2, document.RightMargin - len2, document.PageSize.GetBottom(document.BottomMargin) - 10);
            }

            public override void OnCloseDocument(PdfWriter writer, Document document)
            {
                base.OnCloseDocument(writer, document);
                //BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                //foreach (PdfTemplate item in templates)
                //{
                //    item.BeginText();
                //    item.SetFontAndSize(bf, 10);
                //    item.SetTextMatrix(0, 0);
                //    item.ShowText("" + writer.PageNumber);
                //    item.EndText();
                //}
            }
        }
    }
}
