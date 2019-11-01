<!--<div class="container">
-->
	<div class="panel panel-default">
		<div class="panel-heading" style="font-size: large;">
			<strong class="">Member Sign In</strong>
		</div>
		<div class="panel-body">
			<form class="form-vertical" role="form" method="post" action="/ACE/login/ACELogin.asp">
				<div class="form-group">
					<label for="txtUserName" class="control-label">User ID</label>
					<input type="text" class="form-control" name="txtUserName" id="txtusername" placeholder="User ID" required autofocus>
				</div>
				<div class="form-group">
					<label for="txtPassword" class="control-label">Password</label>
					<input type="password" class="form-control" name="txtPassword" id="txtPassword" placeholder="Password" required>
				</div>
				<div class="checkbox">
					<label class="">
						<input type="checkbox" name="chkRememberMe" value="ON">Remember me
					</label>
				</div>
				<div class="form-group last" style="text-align: center;">
					<button type="submit" class="btn btn-primary btn-md">Login</button>
					<button type="reset" class="btn btn-default btn-md">Reset</button>
				</div>
				<div class="form-group text-center">
					Forgot your password? <a href="password-recovery.aspx">Click here.</a>
				</div>
			</form>	
		</div>
	</div>
<!--	
</div>
-->