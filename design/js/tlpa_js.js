// login modal box
jQuery(document).ready(function(){
	// show login box and mask when login link is clicked
	$("#js-showLoginModal").click(function() {
		$("div#js-tLoginMask").addClass("hideMask");
		$("div#js-loginModal").css("z-index", "100");
		$("div#js-loginModal").addClass("showLoginModal");
		return false;
 	});
	// close button ("x") on login box, clear form
 	$("#js-closeLoginModal").click(function() {
 		$("div#js-loginModal").css("z-index", "-1");
		$("div#js-loginModal").removeClass("showLoginModal");
		$("div#js-tLoginMask").removeClass("hideMask");
		document.getElementById("js-tlpaLoginForm").reset();
		return false;
	});
	// close login box when clicking on mask, clear form
	$("div#js-tLoginMask").click(function() {
 		$("div#js-loginModal").css("z-index", "-1");
		$("div#js-loginModal").removeClass("showLoginModal");
		$("div#js-tLoginMask").removeClass("hideMask");
		document.getElementById("js-tlpaLoginForm").reset();
		return false;
	});
});
// end login modal box

// top of page button
jQuery(document).ready(function() {
	var offset = 100;
	var duration = 500;
	
	jQuery(window).scroll(function() {
		if (jQuery(this).scrollTop() > offset) {
			jQuery('.back-to-top').fadeIn(duration);
		} else {
			jQuery('.back-to-top').fadeOut(duration);
		}
	});
	
	jQuery('.back-to-top').click(function(event) {
		event.preventDefault();
		jQuery('html, body').animate({scrollTop: 0}, duration);
		return false;
	});
});
// end top of page button