<?php
if($_POST)
{
require('captcha/constant.php');
    
     $user_name      = filter_var($_POST["name"], FILTER_SANITIZE_STRING);
    $user_email     = filter_var($_POST["email"], FILTER_SANITIZE_EMAIL);
    $company_name     = filter_var($_POST["company_name"], FILTER_SANITIZE_STRING);
     $comapny_url     = filter_var($_POST["comapny_url"], FILTER_SANITIZE_STRING);
      $message     = filter_var($_POST["message"], FILTER_SANITIZE_STRING);
      
      $response     = filter_var($_POST["response"], FILTER_SANITIZE_STRING);
      
      
    if(empty($user_name)) {
		$empty[] = "<b>Name</b>";		
	}
	if(empty($user_email)) {
		$empty[] = "<b>Email</b>";
	}
	if(empty($company_name)) {
		$empty[] = "<b>Company Name</b>";
	}	
	if(empty($comapny_url)) {
		$empty[] = "<b>Comapny url</b>";
	}
	
	if(!empty($empty)) {
		$output = json_encode(array('type'=>'error', 'text' => implode(", ",$empty) . ' Required!'));
        die($output);
	}
	
	if(!filter_var($user_email, FILTER_VALIDATE_EMAIL)){ //email validation
	    $output = json_encode(array('type'=>'error', 'text' => '<b>'.$user_email.'</b> is an invalid Email, please correct it.'));
		die($output);
	}
    
    	//reCAPTCHA validation
	if (isset($_POST['response'])) {
		
		require('captcha/component/recaptcha/src/autoload.php');		
		
		$recaptcha = new \ReCaptcha\ReCaptcha(SECRET_KEY);

		$resp = $recaptcha->verify($_POST['response'], $_SERVER['REMOTE_ADDR']);

		  if (!$resp->isSuccess()) {
				$output = json_encode(array('type'=>'error', 'text' => '<b>Captcha</b> Validation Required!'));
				die($output);				
		  }	
	}
    
    
    
    
    	$toEmail = "info@digitcomindia.com,chaudhary@digitalsparx.com";
	$mailHeaders = "From: " . $user_name . "<" . $user_email . ">\r\n";
	$mailBody = "User Name: " . $user_name . "\n";
	$mailBody .= "User Email: " . $user_email . "\n";
	$mailBody .= "Comapny Name: " . $company_name . "\n";
	$mailBody .= "Message: " . $message . "\n";

	if (mail($toEmail, "Contact Mail", $mailBody, $mailHeaders)) {
	    $output = json_encode(array('type'=>'message', 'text' => 'Hi '.$user_name .', thank you for the comments. We will get back to you shortly.'));
	    die($output);
	} else {
	    $output = json_encode(array('type'=>'error', 'text' => 'Unable to send email, please contact'.SENDER_EMAIL));
	    die($output);
	}
   
}
?>