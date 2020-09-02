
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using TCSProject.Models;
using Excel= Microsoft.Office.Interop.Excel;
using System.IO;

namespace TCSProject.Controllers
{
    public class AdminController : Controller
    {

        SqlConnection con = new SqlConnection();
        SqlCommand com = new SqlCommand();
        SqlDataReader dr;
        SqlDataAdapter da;
        DataSet ds = new DataSet();


        // GET: Admin
        public ActionResult Index()
        {


            return View();
        }

        void connectionstring()
        {
            con.ConnectionString = "data source = SURYATEJA\\MS; initial catalog = EMP; user id = sa; password = monuSurya; multipleactiveresultsets = True; application name = EntityFramework";
        }

        [HttpPost]
        public ActionResult VerifyAdmin(Admin adm)
        {
            try
            {


                connectionstring();
                con.Open();
                com.Connection = con;
                com.CommandText = "select * from AdmTable where AdmId='" + adm.AdmId + "' and Password='" + adm.Password + "'";
                dr = com.ExecuteReader();


            }
            catch (Exception ex)
            {
                Response.Write(ex.ToString());
            }


            if (dr.Read() == false)
            {

                adm.LoginErrorMessage = "Invalid Credentials !";
                //return View("Login", adm);
                //Response.Write("<script language='javascript'>alert('Invalid Credentials !');</script>");

                return RedirectToAction("Instructions", "Admin");
            }
            else
            {
                con.Close();
                Session["AdmId"] = adm.AdmId;
                return RedirectToAction("Home", "Admin");

            }

        }

        public ActionResult Logout()
        {
            Session.Abandon();
            return RedirectToAction("Login", "Employee");
        }

        public ActionResult Instructions()
        {
            return View();

        }

        public ActionResult Home()
        {
            connectionstring();
            con.Open();
            da = new SqlDataAdapter("Select * from EmplTable", con);
            da.Fill(ds);

            List<Employee> emp = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)

            {

                emp.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString() });

            }

            ViewData.Model = emp;
            return View();
        }


        public ActionResult SearchForEmployee(int Id)
        {
            ViewBag.Message = Id.ToString();
            //Employee emp = new Employee();
            return View();
        }


        [HttpGet]
        public ActionResult ViewEmployee()
        {

            return View();
        }

        [HttpPost]
        public ActionResult ViewEmployee(Employee emp)
        {

            //connectionstring();
            //con.Open();
            //com.Connection = con;
            //com.CommandText = "select * from AdmTable where AdmId='" + emp.EmpId + "'";
            //dr = com.ExecuteReader();

            connectionstring();
            con.Open();
            da = new SqlDataAdapter("Select * from EmplTable where EmpId='" + emp.EmpId + "'", con);
            da.Fill(ds);
            List<Employee> emp1 = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)

            {

                emp1.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(),  Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString(), AssetId = int.Parse(dr[12].ToString()), TokenId = int.Parse(dr[13].ToString()), Location= dr[15].ToString() });

            }
            if (emp1.Count() == 0)
            {
                ViewBag.IsEmployeePresent = false;
                return View();
            }
            else
            {
                ViewBag.IsEmployeePresent = true;
                ViewBag.EmployeeNumber = emp.EmpId.ToString();
                Session["temp"] = emp.EmpId;
            }

            ViewData.Model = emp1[0];
            return View();

        }


        [HttpGet]
        public ActionResult AddEmployee()
        {
            return View();

        }

        [HttpPost]
        public ActionResult AddEmployee(Employee emp)
        {
            connectionstring();
            con.Open();

            //String MyString;
            //DateTime MyDateTime;
            //MyDateTime = new DateTime();
            //MyString = emp.DOB.Date.ToString();
            //MyDateTime = DateTime.ParseExact(MyString, "yyyy-MM-dd",null);

            string insertQuery = "select count(EmpId) from EmplTable where EmpId='" + emp.EmpId + "'";
            SqlCommand cmd = new SqlCommand(insertQuery, con);
            int count = Convert.ToInt32(cmd.ExecuteScalar().ToString());
            if (count != 0)
            {
                ViewBag.AddEmpStatus = false;
                ViewBag.Message = "There's already an employee with given Employee ID";

            }
            else
            {

                insertQuery = "insert into EmplTable(EmpId,FirstName,LastName,Email,ProjectId,WONNumber, ProjectDetails, AssetId, TokenId, Location) values('" + emp.EmpId
                    + "','" + emp.FirstName + "','" + emp.LastName + "','" + emp.Email + "','" + emp.ProjectId + "','" + emp.WONNumber + "','" + emp.ProjectDetails + "','" + emp.AssetId  + "','" + emp.TokenId + "','" + emp.Location + "');" ;
                cmd = new SqlCommand(insertQuery, con);
                int res = cmd.ExecuteNonQuery();
                if (res != 0)
                {
                    ViewBag.AddEmpStatus = true;
                    ViewBag.Message = "Successfully added employee " + emp.EmpId.ToString() + "!";
                    ModelState.Clear();
                }

            }
            return View();
        }


        [HttpGet]
        public ActionResult DeleteEmployee()
        {
            return View();
        }

        [HttpPost]
        public ActionResult DeleteEmployee(Employee emp)
        {
            connectionstring();
            con.Open();
            string insertQuery = "Delete from EmplTable where EmpId='" + emp.EmpId  + "'";
            SqlCommand cmd = new SqlCommand(insertQuery, con);
            int res = cmd.ExecuteNonQuery();
            if (res != 0)
            {
                ViewBag.DeleteStatus = true;
                ViewBag.Message = "Successfully deleted Employee " + emp.EmpId;
                return View();
            }
            else
            {
                ViewBag.DeleteStatus = false;
                ViewBag.Message = "Could not find an employee with given details. Enter correct details.";
            }
            return View();
        }

        [HttpGet]
        public ActionResult UpdateEmployeeDetails()
        {
            return View();
        }



        [HttpPost]
        public ActionResult UpdateEmployeeDetails(Employee emp)
        {
            connectionstring();
            con.Open();

            da = new SqlDataAdapter("Select * from EmplTable where EmpId='" + emp.EmpId + "'", con);
            da.Fill(ds);
            List<Employee> emp1 = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)
            {

                emp1.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString(), AssetId = int.Parse(dr[12].ToString()), TokenId = int.Parse(dr[13].ToString()), Location = dr[15].ToString() });

            }
            if (emp1.Count() == 0)
            {
                ViewBag.IsEmployeePresent = false;
                ViewBag.Message = "No such employee found.";
                return View();
            }
            else
            {
                ViewBag.IsEmployeePresent = true;
                Session["update"] = emp.EmpId;
                
            }
            
            ViewData.Model = emp1[0];
            return View();
        }


        [HttpPost]
        
        public ActionResult Later(Employee emp)
        {

            connectionstring();
            con.Open();
            emp.EmpId = (int)Session["update"];
            string query = "UPDATE EmplTable SET FirstName = @FN , LastName = @LN, Email = @Em , ProjectId= @PId, WONNumber = @WON, ProjectDetails = @PDt, AssetId = @Ast, TokenId = @Tkn, Location = @Lcn WHERE EmpId= @EID";
            SqlCommand cmd = new SqlCommand(query, con);
            cmd.Parameters.AddWithValue("@FN", emp.FirstName);
            cmd.Parameters.AddWithValue("@LN", emp.LastName);
            cmd.Parameters.AddWithValue("@Em", emp.Email);
            cmd.Parameters.AddWithValue("@PId", emp.ProjectId);
            cmd.Parameters.AddWithValue("@WON", emp.WONNumber);
            cmd.Parameters.AddWithValue("@PDt", emp.ProjectDetails); 
            cmd.Parameters.AddWithValue("@Ast", emp.AssetId);
            cmd.Parameters.AddWithValue("@Tkn", emp.TokenId);
            cmd.Parameters.AddWithValue("@Lcn", emp.Location);
            cmd.Parameters.AddWithValue("@EID", Session["update"]);
            int res=cmd.ExecuteNonQuery();
            if(res!=0)
            {
                
                ViewBag.UpdateEmpStatus = true;
                ViewBag.Message = "Successfully updated details: " + emp.EmpId  ;
            }
            else
            {
                ViewBag.UpdateEmpStatus = false;
                ViewBag.Message = " Something went wrong. Please retry.";
            }
            return View();
        }

        [HttpGet]
        public ActionResult ForgotPassword()
        {

            return View();
        }


        [NonAction]
        public void SendVerificationLinkEmail(string emailID, string activationCode, string emailFor = "VerifyAccount")
        {
            var verifyUrl = "/Admin/" + emailFor + "/" + activationCode;
            var link = Request.Url.AbsoluteUri.Replace(Request.Url.PathAndQuery, verifyUrl);

            var fromEmail = new MailAddress("fanofmyplayer@gmail.com", "Employee Management Portal");
            var toEmail = new MailAddress(emailID);
            var fromEmailPassword = "monuSurya"; // Replace with actual password

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
            

            connectionstring();
            con.Open();
            string insertQuery = "select count(AdmId) from AdmTable where EmailId='" + EmailId + "'";
            SqlCommand cmd = new SqlCommand(insertQuery, con);
            int count = Convert.ToInt32(cmd.ExecuteScalar().ToString());
            if (count > 0)
            {
                string resetCode = Guid.NewGuid().ToString();
                SendVerificationLinkEmail(EmailId, resetCode, "ResetPassword");
                insertQuery = "Update AdmTable SET ResetPasswordCode= @RPC where EmailId= @mail";
                cmd = new SqlCommand(insertQuery, con);
                cmd.Parameters.AddWithValue("@RPC", resetCode);
                cmd.Parameters.AddWithValue("@mail", EmailId);
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
            string insertQuery = "select count(AdmId) from AdmTable where ResetPasswordCode='" + id + "'";
            SqlCommand cmd = new SqlCommand(insertQuery, con);
            int count = Convert.ToInt32(cmd.ExecuteScalar().ToString());
            if (count > 0)
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
            var message = "";
            if(model.NewPassword !=  model.ConfirmPassword)
            {
                ViewBag.ResetStatus = false;
                ViewBag.message = "Passwords do not match !";
                return View(model);
            }
            connectionstring();
            con.Open();
            string insertQuery = "select count(AdmId) from AdmTable where ResetPasswordCode='" + model.ResetCode + "'";
            SqlCommand cmd = new SqlCommand(insertQuery, con);
            int count = Convert.ToInt32(cmd.ExecuteScalar().ToString());

            if (count > 0)
            {
                message = "";

                insertQuery = "select AdmId from AdmTable where ResetPasswordCode= '" + model.ResetCode + "'";
                SqlCommand cmd2 = new SqlCommand(insertQuery, con);
                string eid = cmd2.ExecuteScalar().ToString();


                insertQuery = "Update AdmTable SET Password= @pwd , ResetPasswordCode= @rpc where AdmId= @eid";
                SqlCommand cmd1 = new SqlCommand(insertQuery, con);
                cmd1.Parameters.AddWithValue("@pwd", model.NewPassword);
                cmd1.Parameters.AddWithValue("@rpc", "");
                cmd1.Parameters.AddWithValue("@eid", eid);

                int res = cmd1.ExecuteNonQuery();
                
                    message = "New password updated successfully";
                ViewBag.ResetStatus = true;
                ViewBag.message = message;
                return View(model);
            }


            message = "Something invalid";

            ViewBag.Message = message;
            return View(model);
        }

        
        public ActionResult Report()
        {
            List<SelectListItem> MyLocations = new List<SelectListItem>() {
            new SelectListItem {
                Text = "Hyderabad", Value = "Hyderabad"
            },
            new SelectListItem {
                Text = "Mumbai", Value = "Mumbai"
            },
            new SelectListItem {
                Text = "Pune", Value = "Pune"
            },
            new SelectListItem {
                Text = "Chennai", Value = "Chennai"
            },
            new SelectListItem {
                Text = "Bangalore", Value = "Bangalore"
            },
            new SelectListItem {
                Text = "Delhi", Value = "Delhi"
            },
        };



            List<SelectListItem> MySkills = new List<SelectListItem>() {
            new SelectListItem {
                Text = "Web", Value = "Web"
            },
            new SelectListItem {
                Text = "Python", Value = "Python"
            },
            new SelectListItem {
                Text = "DBMS", Value = "DBMS"
            },
            new SelectListItem {
                Text = "Android", Value = "Android"
            },
            new SelectListItem {
                Text = "Machine Learning", Value = "Machine Learning"
            },
            new SelectListItem {
                Text = "AI", Value = "AI"
            },
        };


            ViewBag.locations = MyLocations;
            ViewBag.skills = MySkills;

            return View();
        }



        [HttpPost]
        public ActionResult LocationWiseReport(FormCollection form)
        {
            string strlocations = form["locations"].ToString();
            Session["Location"] = strlocations;
            ViewBag.SelectedLocation = strlocations;
            connectionstring();
            con.Open();
            da = new SqlDataAdapter("Select * from EmplTable where Location='"+ strlocations +"'", con);
            da.Fill(ds);

            List<Employee> emp = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)

            {

                emp.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString() });

            }

            ViewData.Model = emp;
            return View();
            
        }


        public ActionResult LWR()
        {
            connectionstring();
            con.Open();
            string strlocations = Session["Location"].ToString();
            da = new SqlDataAdapter("Select * from EmplTable where Location='" + strlocations + "'", con);
            da.Fill(ds);

            List<Employee> emp = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)

            {

                emp.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString() });

            }


            Microsoft.Office.Interop.Excel.Application excelfile = new Microsoft.Office.Interop.Excel.Application();
            excelfile.Application.Workbooks.Add(Type.Missing);
            excelfile.Cells[1, 1] = "Employee ID";
            excelfile.Cells[1, 2] = "First Name";
            excelfile.Cells[1, 3] = "Last Name";
            excelfile.Cells[1, 4] = "Email";
            excelfile.Cells[1, 5] = "Project ID";
            excelfile.Cells[1, 6] = "WON Number";
            excelfile.Cells[1, 7] = "Project Details";

            int i = 2;
            foreach (Employee temp in emp)
            {
                excelfile.Cells[i, 1] = temp.EmpId.ToString();
                excelfile.Cells[i, 2] = temp.FirstName.ToString();
                excelfile.Cells[i, 3] = temp.LastName.ToString();
                excelfile.Cells[i, 4] = temp.Email.ToString();
                excelfile.Cells[i, 5] = temp.ProjectId.ToString();
                excelfile.Cells[i, 6] = temp.WONNumber.ToString();
                excelfile.Cells[i, 7] = temp.ProjectDetails.ToString();

                i += 1;
            }

            excelfile.Columns.AutoFit();
            excelfile.Visible = true;
            return View("ViewEmployee");
        }
        [HttpPost]
        public ActionResult SkillWiseReport(FormCollection form)
        {
            string strskills = form["skills"].ToString();
            Session["Skills"] = strskills;
            ViewBag.Selectedskill = strskills;
            connectionstring();
            con.Open();
            da = new SqlDataAdapter("Select E.* from EmplTable E, Skills S where E.EmpId= S.EmpId and S.Skill='" + strskills+ "'", con);
            da.Fill(ds);

            List<Employee> emp = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)

            {

                emp.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString() });

            }

            ViewData.Model = emp;

            return View();
        }

        public ActionResult SWR()
        {
            connectionstring();
            con.Open();
            string strlocations = Session["Skills"].ToString();
            da = new SqlDataAdapter("Select E.* from EmplTable E, Skills S where E.EmpId= S.EmpId and S.Skill='" + strlocations + "'", con);
            da.Fill(ds);

            List<Employee> emp = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)

            {

                emp.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString() });

            }


            Microsoft.Office.Interop.Excel.Application excelfile = new Microsoft.Office.Interop.Excel.Application();
            excelfile.Application.Workbooks.Add(Type.Missing);
            excelfile.Cells[1, 1] = "Employee ID";
            excelfile.Cells[1, 2] = "First Name";
            excelfile.Cells[1, 3] = "Last Name";
            excelfile.Cells[1, 4] = "Email";
            excelfile.Cells[1, 5] = "Project ID";
            excelfile.Cells[1, 6] = "WON Number";
            excelfile.Cells[1, 7] = "Project Details";

            int i = 2;
            foreach (Employee temp in emp)
            {
                excelfile.Cells[i, 1] = temp.EmpId.ToString();
                excelfile.Cells[i, 2] = temp.FirstName.ToString();
                excelfile.Cells[i, 3] = temp.LastName.ToString();
                excelfile.Cells[i, 4] = temp.Email.ToString();
                excelfile.Cells[i, 5] = temp.ProjectId.ToString();
                excelfile.Cells[i, 6] = temp.WONNumber.ToString();
                excelfile.Cells[i, 7] = temp.ProjectDetails.ToString();

                i += 1;
            }

            excelfile.Columns.AutoFit();
            excelfile.Visible = true;
            return View("ViewEmployee");
        }

        [HttpPost]
        public ActionResult TeamWiseReport(FormCollection form)
        {

            string TeamId = form["TeamId"].ToString();
            ViewBag.SelectedTeam = TeamId;
            Session["Team"] = TeamId;
            connectionstring();
            con.Open();
            da = new SqlDataAdapter("Select * from EmplTable where ProjectId='" + TeamId + "'", con);
            da.Fill(ds);

            List<Employee> emp = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)

            {

                emp.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString() });

            }

            ViewData.Model = emp;
            return View();
        }


        public ActionResult TWR()
        {
            connectionstring();
            con.Open();
            string strlocations = Session["Team"].ToString();
            da = new SqlDataAdapter("Select * from EmplTable where ProjectId='" + strlocations + "'", con);
            da.Fill(ds);

            List<Employee> emp = new List<Employee>();

            foreach (DataRow dr in ds.Tables[0].Rows)

            {

                emp.Add(new Employee() { EmpId = int.Parse(dr[0].ToString()), FirstName = dr[1].ToString(), LastName = dr[2].ToString(), Email = dr[6].ToString(), ProjectId = int.Parse(dr[8].ToString()), WONNumber = int.Parse(dr[9].ToString()), ProjectDetails = dr[10].ToString() });

            }


            Microsoft.Office.Interop.Excel.Application excelfile = new Microsoft.Office.Interop.Excel.Application();
            excelfile.Application.Workbooks.Add(Type.Missing);
            excelfile.Cells[1, 1] = "Employee ID";
            excelfile.Cells[1, 2] = "First Name";
            excelfile.Cells[1, 3] = "Last Name";
            excelfile.Cells[1, 4] = "Email";
            excelfile.Cells[1, 5] = "Project ID";
            excelfile.Cells[1, 6] = "WON Number";
            excelfile.Cells[1, 7] = "Project Details";

            int i = 2;
            foreach (Employee temp in emp)
            {
                excelfile.Cells[i, 1] = temp.EmpId.ToString();
                excelfile.Cells[i, 2] = temp.FirstName.ToString();
                excelfile.Cells[i, 3] = temp.LastName.ToString();
                excelfile.Cells[i, 4] = temp.Email.ToString();
                excelfile.Cells[i, 5] = temp.ProjectId.ToString();
                excelfile.Cells[i, 6] = temp.WONNumber.ToString();
                excelfile.Cells[i, 7] = temp.ProjectDetails.ToString();

                i += 1;
            }

            excelfile.Columns.AutoFit();
            excelfile.Visible = true;
            return View("ViewEmployee");
        }


        [HttpGet]
        public ActionResult Upload()
        {

            return View();
        }

        


        public ActionResult Success()
        {
            return View();
        }



        public ActionResult Export()
        {

            connectionstring();
            con.Open();
            int EmpId = int.Parse( Session["temp"].ToString() );
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
            excelfile.Cells[10, 1] = "Token ID";
            excelfile.Cells[10, 2] = emp1[0].TokenId;
            excelfile.Cells[11, 1] = "Location";
            excelfile.Cells[11, 2] = emp1[0].Location;


            //excelfile.Columns.AutoFit();
            excelfile.Visible = true;

            return View("ViewEmployee");
        }



        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase excelfile)
        {
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select an ecxel file";
                return View();
            }
            else
            {
                if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/ExcelFiles/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);

                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<Employee> emp = new List<Employee>();

                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        Employee e = new Employee();
                        e.EmpId = int.Parse(((Excel.Range)range.Cells[row, 1]).Text);
                        e.FirstName = ((Excel.Range)range.Cells[row, 2]).Text;
                        e.LastName = ((Excel.Range)range.Cells[row, 3]).Text;

                        e.Email = ((Excel.Range)range.Cells[row, 4]).Text;
                        e.ProjectId = int.Parse(((Excel.Range)range.Cells[row, 5]).Text);
                        e.WONNumber = int.Parse(((Excel.Range)range.Cells[row, 6]).Text);
                        e.ProjectDetails = ((Excel.Range)range.Cells[row, 7]).Text;
                        e.AssetId = int.Parse(((Excel.Range)range.Cells[row, 8]).Text);
                        e.TokenId = int.Parse(((Excel.Range)range.Cells[row, 9]).Text);
                        e.Location = ((Excel.Range)range.Cells[row, 10]).Text;
                        e.LanId = int.Parse(((Excel.Range)range.Cells[row, 11]).Text);

                        emp.Add(e);
                    }

                    ViewBag.emplo = emp;

                    connectionstring();
                    con.Open();


                    List<Employee> AddedEmp = new List<Employee>();
                    List<Employee> Existing = new List<Employee>();
                    foreach (var e in emp)
                    {
                        string insertQuery = "select count(EmpId) from EmplTable where EmpId='" + e.EmpId + "'";
                        SqlCommand cmd = new SqlCommand(insertQuery, con);
                        int count = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                        if (count > 0)
                        {
                            Existing.Add(e);
                        }
                        else
                        {
                            insertQuery = "insert into EmplTable(EmpId,FirstName,LastName,Email,ProjectId,WONNumber, ProjectDetails, AssetId, TokenId, Location) values('" + e.EmpId
                    + "','" + e.FirstName + "','" + e.LastName + "','" + e.Email + "','" + e.ProjectId + "','" + e.WONNumber + "','" + e.ProjectDetails + "','" + e.AssetId + "','" + e.TokenId + "','" + e.Location + "');";
                            cmd = new SqlCommand(insertQuery, con);
                            int res = cmd.ExecuteNonQuery();

                            AddedEmp.Add(e);

                        }


                    }
                    ViewBag.added = AddedEmp;
                    ViewBag.exist = Existing;

                    return View("Success");

                }
                else
                {
                    ViewBag.Error = "The selected file is not excel";
                    return View();
                }

            }

        }

    }

}
        

 
