
<?php
require('captcha/constant.php');
?>
<!DOCTYPE html>
<html>

<head>
<meta charset="UTF-8">
<title>:: Home | Digitcom India Technologies ::</title>
<meta name="description" content="">
<meta name="keywords" content="">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
  <link rel="icon" href="https://www.w3schools.com/tags/demo_icon.gif" type="image/gif" sizes="16x16">

<link href='http://fonts.googleapis.com/css?family=Signika:600,400,300' rel='stylesheet' type='text/css'>
<link href="style.css" rel="stylesheet" type="text/css" media="screen">
<link href="style-headers.css" rel="stylesheet" type="text/css" media="screen">
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>

<style>.error{background-color: #F7902D;  margin-bottom: 40px;}
	 .success {
    background-color: #48e0a4;
    color: white;
    width: 47%;
    padding: 1%;
}

</style>
<script>
	$(document).ready(function (e){
		$("#frmContact").on('submit',(function(e){
			e.preventDefault();
			
			/*$("#mail-status").hide();
			$('#send-message').hide();
			$('#loader-icon').show();*/
			$.ajax({
			    url:'contact_submit.php',
			    type:'post',
			    data:{
			       name:$('input[name="name"]').val(),
			       	email:$('input[name="email"]').val(),
				   company_name:$('input[name="company_name"]').val(),
					comapny_url:$('input[name="comapny_url"]').val(),
					
					message:$('textarea[id="message"]').val(),
					response:$('textarea[id="g-recaptcha-response"]').val(),
				
				
			    },
			    success:function(data){
			        
			        
			        var Obj=JSON.parse(data);
			        console.log(Obj);
			     if(Obj.type="message"){
			         
			         	$("#status").attr("class","success");
			         	
			     }else if(Obj.type="error"){
			         
			         $("#status").attr("class","error");
			     }
			     
			     
			     $("#status").html(Obj.text);
			       
			        
			    }
			    
			    
			})
		}));
	});
	</script>

	<script src='https://www.google.com/recaptcha/api.js'></script>	

</head>

<body class="home">
<div class="root">
  <header class="h3 sticky-enabled no-topbar">
    <section class="top">
      <div>
        <p>Call us: +91-9311322320 &nbsp;&nbsp;|&nbsp;&nbsp; E-mail: <a href="mailto:info@digitcomindia.com" style="color:#3f3f3f">info@digitcomindia.com</a></p>
      </div>
    </section>
    <section class="main-header">
     <div class="title one-logo"><a href="index.html"><img src="images/logo.png" alt="logo"></a> </div>
      <nav class="mainmenu">
        <ul>
          <li><a href="index.html">Home</a></li>
          <li><a href="company-profile.html">About US</a>
            <ul>
              <li><a href="company-profile.html">Company Profile</a></li>
            
            </ul>
          </li>
          <li><a href="#">Products & Services</a>
            <ul>
              <li><a href="vsat.html">VSAT</a></li>
              <li><a href="mobile-networks.html">Mobile Networks</a></li>
              <li><a href="fm.html">FM broadcasting</a></li>
              <li><a href="outdoor-media.html">Outdoor Media</a></li>
            </ul>
          </li>
         <li ><a href="gallery.html">Gallery</a></li>
          <li><a href="clients.html">Clients</a></li>
          <li><a href="career.html">Carrier</a></li>
          <li class="current-menu-item"><a href="contact.html">Contact</a></li>
        </ul>
        <select id="sec-nav" name="sec-nav">
          <option>Main Menu</option>
          <option value="index.html">Home</option>
          <option value="landing-page.html">- About Us</option>
          <option value="company-profile.html">- Company Profile</option>
        
          <option>- Products & Services</option>
          <option value="vsat.html">- VSAT</option>
          <option value="mobile-networks.html" selected="selected">- Telecom</option>
          <option value="fm.html">- FM broadcasting</option>
          <option value="outdoor-media.html">- Outdoor Media</option>
          <option value="gallery.html">Gallery</option>
          <option value="career.html">- Career</option>
          <option value="contact.html">- Contact</option>
        </select>
      </nav>
      <div class="clear"></div>
    </section>
  </header>
   <div class="ttm-page-title-row bann6">
            <div class="container">
                <div class="row">
                    <div class="col-md-12"> 
                        <div class="title-box ttm-textcolor-white">
                            <!-- <div class="page-title-heading">
                                <h1 class="title">About Us</h1>
                               
                            </div> --><!-- /.page-title-captions -->
                           <!--  <div class="breadcrumb-wrapper">
                                <div class="container">
                                    <div class="breadcrumb-wrapper-inner">
                                        <span>
                                            <a title="Go to Delmont." href="index.html" class="home">
											&nbsp;&nbsp;Home</a>
                                        </span>
                                        <span class="ttm-bread-sep">&nbsp; | &nbsp;</span>
                                        <span>About Us</span>
                                    </div>
                                </div>
                            </div> -->
                        </div>
                    </div><!-- /.col-md-12 -->  
                </div><!-- /.row -->  
            </div><!-- /.container -->                      
        <div></div></div>
  <section class="breadcrumb p07">
    <p><a href="#">Home</a> Contact Us</p>
  </section>
  <section class="content contact">
    <article class="main">
    <h3 style="border-bottom: 1px solid #dfdfdf;    font-size: 1.5em;    margin-top: 15px;"><span>Get in touch with us</span></h3>
      <form id="frmContact" action="" method="POST" novalidate="novalidate" class="contact-formsss">
        <p class="half">
          <label for="name">Name</label>
          <input required name="name" id="name">
        </p>
        <p class="half">
          <label for="email">E-mail</label>
          <input type="email" required name="email" id="email">
        </p>
        <p>
          <label for="email">Company Name</label>
          <input required name="company_name" id="company_name">
        </p>
        <p>
          <label for="email">Company URL</label>
          <input required name="comapny_url" id="comapny_url">
        </p>
        
        <p>
          <label for="message">Message</label>
          <textarea required name="message" id="message" rows="3" cols="5"></textarea>
        </p>
        
        	<div id="status"></div>
        	<div class="g-recaptcha" data-sitekey="<?php echo SITE_KEY; ?>"></div>	
        <p>
          <button name="send" type="submit" value="1">Send message</button>
        </p>
      </form>
      	
    </article>
    <aside>
      <section>
        <h3><span>Contact details</span></h3>
        <p><strong>Digitcom India Technologies </strong><br>
          01082 ATS GREEN PARADISO,<br>
Sector Chi IV Greater Noida, U.P. 201310 </p>
        <p><strong>Call us:</strong> <br>
+91 9311322320<br>
+91 9811322320 <br>

          <strong>E-mail:</strong> <a href="#">info@digitcomindia.com</a></p>
          <iframe src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d224368.39371590235!2d77.25804244618055!3d28.516983403738145!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x390ce5a43173357b%3A0x37ffce30c87cc03f!2sNoida%2C+Uttar+Pradesh!5e0!3m2!1sen!2sin!4v1536842270100" width="100%" height="250" frameborder="0" style="border:0" allowfullscreen></iframe>
      </section>
      
      
    </aside>
  </section>
     <footer>
    <section class="widgets">
      <article>
        <h3><span>Digitcom India Technologies</span></h3>
        <p>01082, ATS Paradiso<br>
Sector Chi IV Greater Noida<br>
Distt. Gautam Budh Nagar (U.P.), INDIA
      </article>
      <article class="widget_links">
        <h3><span>Quick Links</span></h3>
        <ul>
         
          <li><a href="company-profile.html">About Us</a></li>
         
          <li><a href="clients.html">Clients</a></li>
          <li><a href="career.html">Career</a></li>
          <li><a href="contact.html">Contact</a></li>
        </ul>
      </article>
      <article class="widget_links">
        <h3><span>Products &amp; Services</span></h3>
        <ul>
          <li><a href="vsat.html">VSAT Network Services</a></li>
          <li><a href="mobile-networks.html">Mobile Networks</a></li>
          <li><a href="fm.htmls">FM broadcasting</a></li>
          <li><a href="outdoor-media.html">Outdoor Media</a></li>
        </ul>
      </article>
      <article class="widget_links">
        <h3><span>Connect with us</span></h3>
      <!--   <nav class="social">
          <ul>
            <li><a href="#" class="facebook">Facebook</a></li>
            <li><a href="#" class="twitter">Twitter</a></li>
            <li><a href="#" class="googleplus">Google+</a></li>
            <li><a href="#" class="pinterest">Pinterest</a></li>
            <li><a href="#" class="rss">RSS</a></li>
          </ul>
        </nav> -->
		<span class="phn">
          Phone<br>
          +91 9311322320<br>
          +91 9811322320 <br>
          
          E-mail: <a href="mailto:Info@digitcomindia.com">Info@digitcomindia.com</a></p></span>
      </article>
    </section>
    <section class="bottom">
      <p class="col2">Copyright 2020 Digitcom India | All rights reserved </p>
	  <p class="col2" style="text-align:right;"> Powered by <a target="_blank" href="https://digitalsparx.com">Digital Sparx Technologies </a></p>
    </section>
	
	
  </footer>
</div>
<script type="text/javascript" src="js/scripts.js"></script>
</body>

</html>