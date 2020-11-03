using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using Absher.NotificationService;
using ANS.Client;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using ServiceApplication.Domain.Domains;
using ServiceApplication.Entities;
namespace ServiceApplicationCoreMVC.Areas.Admin.Controllers {
    /// <summary>
    /// Created By Muna Siraj
    /// </summary>

    [Area ("Admin")]
    [Route ("[area]/[controller]")]
    public class SubmitMessageController : Controller {

        private SubmitMessageDomain _SubmitMessageDomain;
        private AbsherSmsNotification _AbsherSmsNotificationService;
        private AbsherResponceSubmitReqDomain _absherResponceSubmitReqDomain;
        Notification Req = new Notification ();
        NotificationRespone Res = new NotificationRespone ();
        RecipientType Recipients = new RecipientType ();
        ParamType PType = new ParamType ();
        AbsherNotificationFault absherNotificationFault = new AbsherNotificationFault ();
        public SubmitMessageController (SubmitMessageDomain SubmitMessageDomain, AbsherSmsNotification AbsherSmsNotificationService, AbsherResponceSubmitReqDomain AbsherResponceSubmitReqDomain) {
            _SubmitMessageDomain = SubmitMessageDomain;
            _AbsherSmsNotificationService = AbsherSmsNotificationService;
            _absherResponceSubmitReqDomain = AbsherResponceSubmitReqDomain;
        }

        public ActionResult Index () {
            var context = new PrincipalContext (ContextType.Domain);
            var principle = UserPrincipal.FindByIdentity (context, HttpContext.User.Identity.Name);
            ViewBag.Name = principle.Name;
            TempData["Name"] = principle.Name;
            return View ();
        }

        [Route ("FindAll")]
        [HttpGet]
        public async Task<ResultList<SubmitMessage>> FindAll () {
            try {
                ResultList<SubmitMessage> result = new ResultList<SubmitMessage> ();
                result = await _SubmitMessageDomain.FindAll ();
                ViewBag.ErrorMessg = "";
                return result;
            } catch (Exception ex) {

                throw new Exception (ex.Message);

            }
        }

        [Route ("FindByID/{Id:int}")]
        [HttpGet]
        public async Task<ResultEntity<SubmitMessage>> FindByID (int id) {
            ResultEntity<SubmitMessage> result = new ResultEntity<SubmitMessage> ();
            result = await _SubmitMessageDomain.FindByID (id);
            return result;
        }

        [Route ("FindByID/{Id:int}")]
        [HttpGet]
        public async Task<ResultEntity<SubmitMessage>> SendCampain (IFormFile file) { }
        public async Task<ResultEntity<NotificationRespone>> Insert (NotificationRespone notificationRespone) {
            ResultEntity<NotificationRespone> result = new ResultEntity<NotificationRespone> ();
            result = await NotificationRespone.Insert (notificationRespone);
            return result
        }
        POST - CREATE[HttpPost]
            [ValidateAntiForgeryToken]
        public async Task<IActionResult> Insert (AbsherResponceSubmitReq AbsherResponceSubmitReq) {
            if (ModelState.IsValid) {
                if valid
                _SubmitMessageDomain.Category.Add (notificationRespone);
                await _absherResponceSubmitReqDomain.Insert (AbsherResponceSubmitReq);
                return RedirectToAction (nameof (Index));
            }
            return View (AbsherResponceSubmitReq);
        }

        [Route ("ImportExcelFile")]
        [HttpPost]

        public async Task<ResultList<NotificationRespone>> ImportExcelFile (string ID, IFormFile formFile, string messageParagraph, string userLoggedInID, string totalVarible)

        {

            ResultList<NotificationRespone> result = new ResultList<NotificationRespone> ();
            if (formFile == null || formFile.Length <= 0) {
                result.Status = 3;
                result.Message = "No data found";
                return result;
            }
            if (!Path.GetExtension (formFile.FileName).Equals (".xlsx", StringComparison.OrdinalIgnoreCase)) {
                result.Status = 1;
                result.Message = "No data found";
                return result;
            }
            HttpResponseMessage res = new HttpResponseMessage ();
            if (res.IsSuccessStatusCode)
                Req.ClientId = "7001044523";
            Req.ClientAuthorization = "kGB4B72DBmAvulfzTFQkbvmHvZlEXYY91wuFBU3v4jE=";
            Req.SubmitId = ID;
            ParamType param1 = new ParamType ();
            ParamType param2 = new ParamType ();
            ParamType param3 = new ParamType ();
            var list = new List<MessageVaribles> ();
            var listabsher = new List<NotificationRespone> ();
            var absherErrorMsgList = new List<AbsherNotificationFault> ();
            using (var stream = new MemoryStream ()) {
                await formFile.CopyToAsync (stream, cancellationToken);
                await formFile.CopyToAsync (stream);
                using (var package = new ExcelPackage (stream)) {
                    var absherErrorMsg = new AbsherNotificationFault ();
                    try {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;
                        int ColCount = worksheet.Dimension.Columns;
                        var RecipientsList = new List<RecipientType> ();
                        for (int row = 2; row <= rowCount; row++) {
                            Recipients = new RecipientType ();
                            Recipients.Language = "AR";
                            Recipients.NationalOrIqamaId = Convert.ToString (worksheet.Cells[row, 1].Value);
                            string rowValue = Convert.ToString (worksheet.Cells[row, 1].Value);
                            if (rowValue.Length < 10) {
                                result.Status = 1;
                                result.Message = "الرجاء التأكد من رقم الهوية أو الاقامة  حيث أنه أقل من 10 أرقام " + rowValue;
                                break;
                            } else {
                                Recipients.NationalOrIqamaId = rowValue;
                                List<ParamType> ParamList = new List<ParamType> ();
                                var allColwithoutFirstCol = ColCount - 1;
                                if (totalVarible == allColwithoutFirstCol.ToString ()) {
                                    if (totalVarible == "0") {
                                        RecipientsList.Add (Recipients);
                                        Req.Recipients = RecipientsList.ToArray ();
                                        var Abshar_result = _AbsherSmsNotificationService.SubmitRequest (Req);
                                    } else {
                                        int i = 1;
                                        ParamType param = new ParamType ();
                                        for (int col = 2; col <= ColCount; col++) {
                                            string h = Convert.ToString (worksheet.Cells[1, col].Value);
                                            param = new ParamType ();
                                            param.Name = "VAR0" + i.ToString ();

                                            string str = Convert.ToString (worksheet.Cells[row, col].Value);

                                            Match match = Regex.Match (str, @"[~`!#$%^&*+=|{}';<>?[\]""]", RegexOptions.IgnoreCase);

                                            if (!match.Success) {

                                                param.Value = str;
                                                ParamList.Add (param);
                                                i++;
                                            } else {
                                                result.Message = "الرجاء التأكد من المحتوى حيث يوجد رمز لايمكن ارساله ( " + match + " )";

                                                result.Status = 1;
                                                return result;
                                            }

                                        }

                                        Recipients.Params = ParamList.ToArray ();
                                        RecipientsList.Add (Recipients);
                                    }
                                } else {

                                    result.Message = allColwithoutFirstCol + "لابد أن يكون عدد المتغييرات";
                                    result.Message = "توجد مشكلة لايمكن الارسال عدد متغييرات الرسالة  غير متناسب مع عدد الأعمدة بالملف المرفق. الرجاء التأكد من أن عدد المتغييرات لابد أن يساوي عدد الأعمدة بالملف المرفق ليتناسب مع قالب الارسال";
                                    result.Status = 1;
                                    return result;

                                }

                            }
                            Req.Recipients = RecipientsList.ToArray ();
                            var Abshar_result = _AbsherSmsNotificationService.SubmitRequest (Req);
                            if (Abshar_result != null) {

                                AbsherResponceSubmitReq resEnity = new AbsherResponceSubmitReq ();
                                resEnity.BatchStatus = Abshar_result.Status.ToString ();
                                resEnity.BatchNumber = Abshar_result.BatchNumber;
                                resEnity.MessageParagraph = messageParagraph;
                                resEnity.SubmitID = ID;

                                resEnity.CreatedBy = HttpContext.User.Identity.Name;

                                await _absherResponceSubmitReqDomain.Insert (resEnity);
                                listabsher.Add (Abshar_result);
                                result.List = listabsher;
                                result.Status = 0;
                                result.Message = "OK";
                                return result;
                            } else {
                                result.Status = 1;
                                result.Message = "Error in Data";
                                return result;
                            }
                        } catch (Exception ex) {
                            if (ex.Message == "Client does not exists") {
                                absherErrorMsg.ErrorCode = "20001";

                                absherErrorMsg.ErrorMessage = "العميل غير مسجل ";
                                result.Message = absherErrorMsg.ErrorMessage + " رقم الخطاء : " + absherErrorMsg.ErrorCode;
                                result.Status = 1;

                            } else if (ex.Message == "Client is Inactive") {
                                absherErrorMsg.ErrorCode = "20002";

                                absherErrorMsg.ErrorMessage = "العميل غير مفعل ";
                                result.Message = absherErrorMsg.ErrorMessage + " رقم الخطاء : " + absherErrorMsg.ErrorCode;
                                result.Status = 1;

                            } else if (ex.Message == "Submit does not exists.") {
                                absherErrorMsg.ErrorCode = "20003";

                                absherErrorMsg.ErrorMessage = "تم اختيار قالب غير صحيح  ";
                                result.Message = absherErrorMsg.ErrorMessage + " رقم الخطاء : " + absherErrorMsg.ErrorCode;
                                result.Status = 1;

                            } else if (ex.Message == "Submit Parameters are  inconsistent with request parameters") {
                                absherErrorMsg.ErrorCode = "10003";

                                absherErrorMsg.ErrorMessage = "تم اختيار قالب غير متوافق  ";
                                result.Message = absherErrorMsg.ErrorMessage + " رقم الخطاء : " + absherErrorMsg.ErrorCode;
                                result.Status = 1;

                            } else if (ex.Message == "Language incorrect") {
                                absherErrorMsg.ErrorCode = "20004";

                                absherErrorMsg.ErrorMessage = "تم اختيار لغة غير صحيحة ";
                                result.Message = absherErrorMsg.ErrorMessage + " رقم الخطاء : " + absherErrorMsg.ErrorCode;
                                result.Status = 1;

                            } else if (ex.Message == "Any of the date is missing or empty") {
                                absherErrorMsg.ErrorCode = "20005";

                                absherErrorMsg.ErrorMessage = "أحد البيانات مفقود أو لايوجد ";
                                result.Message = absherErrorMsg.ErrorMessage + " رقم الخطاء : " + absherErrorMsg.ErrorCode;
                                result.Status = 1;
                                //return result;
                            } else {

                                result.Message = " مشكلة من مزود الخدمة توجد مشكلة لايمكن الارسال عدد متغييرات الرسالة  غير متناسب مع عدد الأعمدة بالملف المرفق. الرجاء التأكد من أن عدد المتغييرات لابد أن يساوي عدد الأعمدة بالملف المرفق ليتناسب مع قالب الارسال، أو الملف الذي تم ارفاقه يحتوي على رقم هوية أو اقامة غير صحيح  ";
                                result.Status = 1;

                            }
                        }
                    }
                }

                return result;
            }
        }

        [Route ("CheckExcelFile")]
        [HttpPost]
        public async Task<ResultList<MessageVaribles>> CheckExcelFile (IFormFile formFile, string totalVarible)

        {

            ResultList<MessageVaribles> result = new ResultList<MessageVaribles> ();
            if (formFile == null || formFile.Length <= 0) {

                result.Status = 3;
                result.Message = "No data found";
                return result;
            }
            if (!Path.GetExtension (formFile.FileName).Equals (".xlsx", StringComparison.OrdinalIgnoreCase)) {

                result.Status = 1;
                result.Message = "No data found";
                return result;
            }
            var list = new List<MessageVaribles> ();
            using (var stream = new MemoryStream ()) {

                await formFile.CopyToAsync (stream);
                using (var package = new ExcelPackage (stream)) {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    if (worksheet.Dimension == null) {
                        result.Status = 1;
                        result.Message = "لا يحتوي الملف على  أرقام هويات الرجاء اضافة هوية أو اقامة على الاقل ";

                    } else {
                        int rowCount = worksheet.Dimension.Rows;
                        if (rowCount == 1) {
                            result.Status = 1;
                            result.Message = "لا يحتوي الملف على  أرقام هويات الرجاء اضافة هوية أو اقامة على الاقل  ";

                        } else {
                            int ColCount = worksheet.Dimension.Columns;
                            int totalCol = 0;
                            for (int row = 2; row <= rowCount; row++) {
                                var allColwithoutFirstCol = ColCount - 1;
                                for (int col = 2; col <= ColCount; col++) {
                                    string dataInCol = Convert.ToString (worksheet.Cells[1, col].Value);

                                    totalCol++;
                                }

                                if (totalVarible == allColwithoutFirstCol.ToString ()) {

                                    var NodataMuiltColumnsList = new List<string> ();

                                    NodataMuiltColumnsList.ToList ().Clear ();
                                    for (int col = 1; col <= ColCount; col++) {
                                        string NodataMuiltColumns = Convert.ToString (worksheet.Cells[row, col].Value);
                                        NodataMuiltColumnsList.Add (NodataMuiltColumns);

                                    }
                                    if (NodataMuiltColumnsList != null && NodataMuiltColumnsList.Any (i => i == "")) {
                                        result.List.Clear ();
                                        result.Message = " يحتوي الملف على عناوين الأعمدة وبعض الحقول الفارغة الرجاء التأكد من وجود بيانات أو حذف الصفوف التي تحتوي على بيانات فارعة واعادة رفع الملف  ";
                                        result.Status = 1;

                                        return result;
                                    } else {

                                        list.Add (new MessageVaribles {
                                            NationalOrIqamaID = worksheet.Cells[row, 1].Value.ToString ().Trim ()
                                        });
                                        result.List = list;
                                        result.Status = 0;

                                    }

                                } else {
                                    result.Message = "توجد مشكلة لايمكن الارسال عدد متغييرات الرسالة  غير متناسب مع عدد الأعمدة بالملف المرفق. الرجاء التأكد من أن عدد المتغييرات لابد أن يساوي عدد الأعمدة بالملف المرفق ليتناسب مع قالب الارسال";
                                    result.Status = 1;

                                }

                            }
                        }
                    }
                }
            }
            return result;
        }

    }
}