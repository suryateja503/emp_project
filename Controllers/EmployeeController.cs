using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using TCSProject.Models;

namespace TCSProject.Controllers
{
    public class EmployeeController : Controller
    {

        Employee emp = new Employee();
        SqlConnection con= new SqlConnection();
        SqlCommand com = new SqlCommand();
        SqlDataAdapter da;
        SqlDataReader dr;
        DataSet ds = new DataSet();
        // GET: Login


        
        [HttpGet]
        public ActionResult Login()
        {
            return View();
        }

        void connectionstring()
        {
            con.ConnectionString = "data source = SURYATEJA\\MS; initial catalog = EMP; user id = sa; password = monuSurya; multipleactiveresultsets = True; application name = EntityFramework";
        }


        [HttpPost]
        public ActionResult Home(Employee emp)
        {
            connectionstring();
            con.Open();
            da = new SqlDataAdapter("Select * from EmplTable where EmpId='" + emp.EmpId + "' and password='" + emp.Password + "'" , con);
            da.Fill(ds);
            List<Employee> emp1 = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)

            {

                emp1.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString() });

            }
            


            if (emp1.Count() ==0)
            {
         
                emp.LoginErrorMessage = "Invalid Credentials !";
                return View("Login", emp);
            }
            else
            {
                con.Close();
                Session["EmpId"] = emp.EmpId;
                ViewBag.PorjectId = emp1[0].ProjectId;
                ViewBag.WONNumber = emp1[0].WONNumber;
                Employee temp = new Employee();
                temp = emp1[0];
                ViewData.Model = temp;
                

                emp.EmpId = emp1[0].EmpId;
                emp.FirstName = emp1[0].FirstName;
                emp.LastName = emp1[0].LastName;
                emp.Email = emp1[0].Email;

                return View();
            }
        }



    
        [HttpGet]
        public ActionResult Home()
        {
            
            return View();
        }
        public ActionResult Logout()
        {
            Session.Abandon();
            return RedirectToAction("Login", "Employee");
        }
        

        [HttpGet]
        public ActionResult ForgotPassword()
        {

            return View();
        }


        [NonAction]
        public void SendVerificationLinkEmail(string emailID, string activationCode, string emailFor = "VerifyAccount")
        {
            var verifyUrl = "/Employee/" + emailFor + "/" + activationCode;
            var link = Request.Url.AbsoluteUri.Replace(Request.Url.PathAndQuery, verifyUrl);

            var fromEmail = new MailAddress("fanofmyplayer@gmail.com", "Employee Management Portal");
            var toEmail = new MailAddress(emailID);
            var fromEmailPassword = "#monuSurya05"; 

            string subject = "";
            string body = "";
            if (emailFor == "VerifyAccount")
            {
                subject = "Your account is successfully created!";
                body = "<br/><br/>We are excited to tell you that your Dotnet Awesome account is" +
                    " successfully created. Please click on the below link to verify your account" +
                    " <br/><br/><a href='" + link + "'>" + link + "</a> ";
            }
            else if (emailFor == "ResetPassword")
            {
                subject = "Reset Password";
                body = "Hi,<br/>We got request for reset your account password. Please click on the below link to reset your password" +
                    "<br/><br/><a href=" + link + ">Reset Password link</a>";
            }


            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(fromEmail.Address, fromEmailPassword)
            };

            using (var message = new MailMessage(fromEmail, toEmail)
            {
                Subject = subject,
                Body = body,
                IsBodyHtml = true
            })
                smtp.Send(message);
        }

        [HttpPost]
        public ActionResult ForgotPassword(string EmailId)
        {
            string message = "";
            bool status = false;

            connectionstring();
            con.Open();
            string insertQuery = "select count(EmpId) from EmplTable where Email='" + EmailId + "'";
            SqlCommand cmd = new SqlCommand(insertQuery, con);
            int count = Convert.ToInt32(cmd.ExecuteScalar().ToString());
            if (count > 0)
            {
                string resetCode = Guid.NewGuid().ToString();
                SendVerificationLinkEmail(EmailId, resetCode, "ResetPassword");
                insertQuery = "Update EmplTable SET ResetPasswordCode= @RPC where Email= @mail";
                cmd = new SqlCommand(insertQuery, con);
                cmd.Parameters.AddWithValue("@RPC", resetCode);
                cmd.Parameters.AddWithValue("@mail", EmailId );
                count = cmd.ExecuteNonQuery();
                message = "Reset password link has been sent to your email id.";
            }
            else
            {
                message = "Account not found";
            }
            ViewBag.Message = message;
      
            return View();


        }



        public ActionResult ResetPassword(string id)
        {
            //Verify the reset password link
            //Find account associated with this link
            //redirect to reset password page
            if (string.IsNullOrWhiteSpace(id))
            {
                return HttpNotFound();
            }

            connectionstring();
            con.Open();
            string insertQuery = "select count(EmpId) from EmplTable where ResetPasswordCode='" + id + "'";
            SqlCommand cmd = new SqlCommand(insertQuery, con);
            int count = Convert.ToInt32(cmd.ExecuteScalar().ToString());
            if(count>0)
            {
                
                    ResetPasswordModel model = new ResetPasswordModel();
                    model.ResetCode = id;
                    return View(model);
                
                
            }
            else
            {
                return HttpNotFound();
            }
        }





        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult ResetPassword(ResetPasswordModel model)
        {
            var message = "Entered Action.... ";
            
                connectionstring();
                con.Open();
                string insertQuery = "select count(EmpId) from EmplTable where ResetPasswordCode='" + model.ResetCode + "'";
                SqlCommand cmd = new SqlCommand(insertQuery, con);
                int count = Convert.ToInt32(cmd.ExecuteScalar().ToString());

                if (count > 0)
                {
                    message = "Employee exists ";

                    insertQuery = "select EmpId from EmplTable where ResetPasswordCode= '" + model.ResetCode + "'";
                    SqlCommand cmd2 = new SqlCommand(insertQuery, con);
                    string eid = cmd2.ExecuteScalar().ToString();


                    insertQuery = "Update EmplTable SET Password= @pwd , ResetPasswordCode= @rpc where EmpId= @eid";
                    SqlCommand cmd1 = new SqlCommand(insertQuery, con);
                    cmd1.Parameters.AddWithValue("@pwd", model.NewPassword);
                    cmd1.Parameters.AddWithValue("@rpc", "");
                    cmd1.Parameters.AddWithValue("@eid", eid);

                    int res = cmd1.ExecuteNonQuery();
                    if (res > 0)
                        message += "New password updated successfully";
                ViewBag.message = message;
                return View(model);
                }

            
                message = "Something invalid";
            
            ViewBag.Message = message;
            return View(model);
        }

        [HttpGet]
        public ActionResult EditDetails()
        {
            connectionstring();
            con.Open();
            int EmpId = (int)Session["EmpId"];
            da = new SqlDataAdapter("Select * from EmplTable where EmpId='" + EmpId + "'", con);
            da.Fill(ds);
            List<Employee> emp1 = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)
            {

                emp1.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString(), AssetId = int.Parse(dr[12].ToString()), TokenId = int.Parse(dr[13].ToString()), Location = dr[15].ToString() });

            }
            

            ViewData.Model = emp1[0];
            return View();
            
        }

        [HttpPost]
        public ActionResult EditDetails(Employee emp)
        {
            connectionstring();
            con.Open();
            emp.EmpId = (int)Session["EmpId"];
            string query = "UPDATE EmplTable SET ProjectId= @PId, WONNumber = @WON, ProjectDetails = @PDt, AssetId = @Ast, TokenId = @Tkn, Location = @Lcn WHERE EmpId= @EID";
            SqlCommand cmd = new SqlCommand(query, con);
           
            cmd.Parameters.AddWithValue("@PId", emp.ProjectId);
            cmd.Parameters.AddWithValue("@WON", emp.WONNumber);
            cmd.Parameters.AddWithValue("@PDt", emp.ProjectDetails);
            cmd.Parameters.AddWithValue("@Ast", emp.AssetId);
            cmd.Parameters.AddWithValue("@Tkn", emp.TokenId);
            cmd.Parameters.AddWithValue("@Lcn", emp.Location);
            cmd.Parameters.AddWithValue("@EID", emp.EmpId);
            int res = cmd.ExecuteNonQuery();
            if (res != 0)
            {

                ViewBag.UpdateEmpStatus = true;
                ViewBag.Message = "Successfully updated details: " + emp.EmpId;
            }
            else
            {
                ViewBag.UpdateEmpStatus = false;
                ViewBag.Message = " Something went wrong. Please retry.";
            }
            return View();
        }

       
        public ActionResult Export()
        {

            connectionstring();
            con.Open();
            int EmpId = (int)Session["EmpId"];
            da = new SqlDataAdapter("Select * from EmplTable where EmpId='" + EmpId + "'", con);
            da.Fill(ds);
            List<Employee> emp1 = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)
            {

                emp1.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString(), AssetId = int.Parse(dr[12].ToString()), TokenId = int.Parse(dr[13].ToString()), Location = dr[15].ToString() });

            }


            Microsoft.Office.Interop.Excel.Application excelfile = new Microsoft.Office.Interop.Excel.Application();
            excelfile.Application.Workbooks.Add(Type.Missing);
            excelfile.Cells[1, 1] = "Employee";
            excelfile.Cells[1, 2] = "Details";


            excelfile.Cells[2, 1] = "Employee ID";
            excelfile.Cells[2, 2] = emp1[0].EmpId;
            excelfile.Cells[3, 1] = "First Name";
            excelfile.Cells[3, 2] = emp1[0].FirstName;
            excelfile.Cells[4, 1] = "Last Name";
            excelfile.Cells[4, 2] = emp1[0].LastName;
            excelfile.Cells[5, 1] = "Email";
            excelfile.Cells[5, 2] = emp1[0].Email;
            excelfile.Cells[6, 1] = "Project ID";
            excelfile.Cells[6, 2] = emp1[0].ProjectId;
            excelfile.Cells[7, 1] = "WON Number";
            excelfile.Cells[7, 2] = emp1[0].WONNumber;
            excelfile.Cells[8, 1] = "Project Details";
            excelfile.Cells[8, 2] = emp1[0].ProjectDetails;
            excelfile.Cells[9, 1] = "Asset ID";
            excelfile.Cells[9, 2] = emp1[0].AssetId;
            excelfile.Cells[10,1] = "Token ID";
            excelfile.Cells[10,2] = emp1[0].TokenId;
            excelfile.Cells[11, 1] = "Location";
            excelfile.Cells[11, 2] = emp1[0].Location;


            excelfile.Columns.AutoFit();
            excelfile.Visible = true;

            return View();
        }

    }
}
