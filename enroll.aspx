<!DOCTYPE html>
<html lang="en">
<head>
	<!-- #include file="includes/default-meta-tags.inc" -->

	<title>ACE | Desk Review of Auto and Property Claims | Enroll</title>

	<!-- #include file="includes/jquery.inc" -->
	<!-- #include file="includes/bootstrap.inc" -->
	<!-- #include file="includes/ie-hacks.inc" -->

	<link href="css/styles.css" rel="stylesheet" />
</head>
<body>
    <script>
        $(function() {
            $("#txtCoName").focus();
        });
    </script>

	<div class="container">
		<!-- #include file="includes/base-content.inc" -->
		<div class="row">
		<!-- #include file="includes/topNavbar.inc" -->
		</div>
		<div class="row">
		<div class="container col-md-2" style="padding-bottom: 15px">
			<!-- #include file="includes/sidebarNav.inc" -->
		</div>
		
		<div class='col-md-10'>
        <div class="well">
            <form class="form-horizontal" method="post" action="/ACE/login/ACELogin.asp?txtState=2">
                <fieldset>
                    <legend>
                        Member Signup
                        <br/>
                        <span style="font-size: 14px; color: gray">An asterisk (<span style="color: red">*</span>) denotes a required field</span>
                    </legend>
                    <div class="row text-center">
                        <h4>Company Information</h4>
                    </div>
                    <br/>
                    <div class="form-group">
                        <label for="txtCoName" class="col-md-2 control-label">Company<span style="color: red">*</span></label>
                        <div class="col-md-10">
                            <input type="text" class="form-control" name="txtCoName" id="txtCoName" required>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtAgencyName" class="col-md-2 control-label">Agency</label>
                        <div class="col-md-10">
                            <input type="text" class="form-control" name="txtAgencyName" id="txtAgencyName">
                            <span class="help-block">Type the name of your agency, if applicable.</span>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtCoAddress" class="col-md-2 control-label">Street Address or PO Box</label>
                        <div class="col-md-10">
                            <input type="text" class="form-control" name="txtCoAddress" id="txtCoAddress">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtCoCity" class="col-md-2 control-label">City<span style="color: red">*</span></label>
                        <div class="col-md-10">
                            <input type="text" class="form-control" name="txtCoCity" id="txtCoCity" required>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="cmbCoState" class="col-md-2 control-label">State/Province<span style="color: red">*</span></label>
                        <div class="col-md-2">
                            <select class="form-control" name="cmbCoState" id="cmbCoState" required>
                                <option></option>
                                <option>AB</option>
                                <option>AK</option>
                                <option>AL</option>
                                <option>AR</option>
                                <option>AZ</option>
                                <option>BC</option>
                                <option>CA</option>
                                <option>CO</option>
                                <option>CT</option>
                                <option>DC</option>
                                <option>DE</option>
                                <option>FL</option>
                                <option>GA</option>
                                <option>HI</option>
                                <option>IA</option>
                                <option>ID</option>
                                <option>IL</option>
                                <option>IN</option>
                                <option>KS</option>
                                <option>KY</option>
                                <option>LA</option>
                                <option>MA</option>
                                <option>MB</option>
                                <option>MD</option>
                                <option>ME</option>
                                <option>MI</option>
                                <option>MN</option>
                                <option>MO</option>
                                <option>MS</option>
                                <option>MT</option>
                                <option>NB</option>
                                <option>NC</option>
                                <option>ND</option>
                                <option>NE</option>
                                <option>NF</option>
                                <option>NH</option>
                                <option>NJ</option>
                                <option>NM</option>
                                <option>NS</option>
                                <option>NT</option>
                                <option>NU</option>
                                <option>NV</option>
                                <option>NY</option>
                                <option>OH</option>
                                <option>OK</option>
                                <option>ON</option>
                                <option>OR</option>
                                <option>PA</option>
                                <option>PE</option>
                                <option>QC</option>
                                <option>RI</option>
                                <option>SC</option>
                                <option>SD</option>
                                <option>SK</option>
                                <option>TN</option>
                                <option>TX</option>
                                <option>UT</option>
                                <option>VA</option>
                                <option>VT</option>
                                <option>WA</option>
                                <option>WI</option>
                                <option>WV</option>
                                <option>WY</option>
                                <option>YT</option>
                                <option>OTHER</option>
                            </select>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtCoZip" class="col-md-2 control-label">Zip Code<span style="color: red">*</span></label>
                        <div class="col-md-2">
                            <input type="text" class="form-control" name="txtCoZip" id="txtCoZip" required>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtCoOffice" class="col-md-2 control-label">Office<span style="color: red">*</span></label>
                        <div class="col-md-10">
                            <input type="text" class="form-control" name="txtCoOffice" id="txtCoOffice" required>
                            <span class="help-block">Type the name of your office branch, region, or program.</span>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtPhoneAc" class="col-md-2 control-label">Work Phone<span style="color: red">*</span></label>
                        <div class="col-md-2">
                            <input type="text" class="form-control" name="txtPhoneAc" id="txtPhoneAc" required maxlength="3">
                        </div>
                        <div class="col-md-2">
                            <input type="text" class="form-control" name="txtPhoneExch" id="txtPhoneExch" required maxlength="3">
                        </div>
                        <div class="col-md-3">
                            <input type="text" class="form-control" name="txtPhoneL4" id="txtPhoneL4" required maxlength="4">
                        </div>
                        <label for="txtExtension" class="col-md-1 control-label">Extension</label>
                        <div class="col-md-2">
                            <input type="number" class="form-control" name="txtPhoneExtn" id="txtPhoneExtn">
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtCoFaxAc" class="col-md-2 control-label">Fax</label>
                        <div class="col-md-2">
                            <input type="text" class="form-control" name="txtCoFaxAc" id="txtCoFaxAc" maxlength="3">
                        </div>
                        <div class="col-md-2">
                            <input type="text" class="form-control" name="txtCoFaxExch" id="txtCoFaxExch" maxlength="3">
                        </div>
                        <div class="col-md-3">
                            <input type="text" class="form-control" name="txtCoFaxL4" id="txtCoFaxL4" maxlength="4">
                        </div>
                    </div>
                    <hr>
                    <div class="row text-center">
                        <h4>Personal Information</h4>
                    </div>
                    <br/>
                    <div class="form-group">
                        <label for="txtFirstName" class="col-md-2 control-label">First Name<span style="color: red">*</span></label>
                        <div class="col-md-4">
                            <input type="text" class="form-control" name="txtFirstName" id="txtFirstName" required>
                        </div>
                        <label for="txtLastName" class="col-md-2 control-label">Last Name<span style="color: red">*</span></label>
                        <div class="col-md-4">
                            <input type="text" class="form-control" name="txtLastName" id="txtLastName" required>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="cmbTitle" class="col-md-2 control-label">Title<span style="color: red">*</span></label>
                        <div class="col-md-2">
                            <select class="form-control" name="cmbTitle" id="cmbTitle" required>
                                <option></option>
                                <option>Claim Rep (see only your assignments)</option>
                                <option>Office Manager (see all assignments from your unit or region)</option>
                                <option>Company Manager (see all assignments from your company)</option>
                            </select>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtEmail" class="col-md-2 control-label">E-mail Address<span style="color: red">*</span></label>
                        <div class="col-md-10">
                            <input type="email" class="form-control" name="txtEmail" id="txtEmail" required>
                        </div>
                    </div>
                    <div class="form-group">
                        <label class="col-lg-2 control-label">Sex<span style="color: red">*</span></label>
                        <div class="col-lg-10">
                            <div class="radio">
                                <label>
                                    <input type="radio" name="grpGender" value="1" checked="checked">
                                    Male
                                </label>
                            </div>
                            <div class="radio">
                                <label>
                                    <input type="radio" name="grpGender" value="2">
                                    Female
                                </label>
                            </div>
                        </div>
                    </div>
                    <hr>
                    <div class="row text-center">
                        <h4>Choose Your Login Information</h4>
                    </div>
                    <br/>
                    <div class="form-group">
                        <label for="txtUserName" class="col-md-2 control-label">User Name<span style="color: red">*</span></label>
                        <div class="col-md-10">
                            <input type="text" class="form-control" name="txtUserName" id="txtUserName" required>
                            <span class="help-block">Must be <b>3 - 20 alphanumeric</b> characters in length.</span>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtPassword" class="col-md-2 control-label">Password<span style="color: red">*</span></label>
                        <div class="col-md-10">
                            <input type="password" class="form-control" name="txtPassword" id="txtPassword" required>
                            <span class="help-block">Keep this private!</span>
                        </div>
                    </div>
                    <div class="form-group">
                        <label for="txtPasswordVerify" class="col-md-2 control-label">Re-enter Password<span style="color: red">*</span></label>
                        <div class="col-md-10">
                            <input type="password" class="form-control" name="txtPasswordVerify" id="txtPasswordVerify" required>
                            <span class="help-block">For verification.</span>
                        </div>
                    </div>
                    <hr>
                    <div class="form-group">
                        <div class="col-md-10 col-md-offset-2">
                            <b>Make sure your email address is correct.</b>
                            <br/>
                            You will be notified via email when your account is set up. Each user is validated manually to ensure security. Your account may take up to 24 hours to set up. Thank you.
                        </div>
                    </div>
                    <div class="form-group">
                        <div class="col-md-10 col-md-offset-2">
                            <button type="submit" name="submitup" id="submitup" class="btn btn-primary btn-md">Submit</button>
                            <button type="reset" class="btn btn-default btn-md">Reset</button>
                        </div>
                    </div>
                    <br/>
                    <div class="form-group text-center">
                        <div class="col-md-10 col-md-offset-2">
                            <b>
                                By submitting this completed form you are agreeing to our terms of use.
                                <br/>
                                Please read our <a href="disclaimer.aspx">Terms of Service Agreement</a>.
                            </b>
                        </div>
                    </div>
                </fieldset>
            </form>
        </div>
    </div>
	</div>

	<!-- #include file="includes/footer.inc" -->
</body>
</html>
