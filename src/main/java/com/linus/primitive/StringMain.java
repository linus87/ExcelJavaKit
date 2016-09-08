package com.linus.primitive;

public class StringMain {

	public static void main(String[] args) {
		String format = "hello %s";
		String result = String.format(format, "world!");
		System.out.println(result);
		
		compare();
		
		split();
		
		runRegexp();
		
		char c = '3';
		System.out.println(c);
	}
	
	private static void compare() {
		System.out.println("Compare two string:---------------------------------");
		String scheme = "http";
		String scheme2 = new String("http");
		System.out.println("http" == "http");
		System.out.println("http" == scheme);
		System.out.println(scheme2 == scheme);
		System.out.println(new String("http") == scheme);
		System.out.println(scheme.equalsIgnoreCase(new String("http")));		
	}
	
	private static void split() {
		System.out.println("Split a string:---------------------------------");
		String text = "Draft>Nomination eDM in approve flow>Nomination eDM approved>Seller nomination_Need approve>Promotion Submitted>Promotion Approved>Notification eDM in approve flow>Notification eDM approved>Seller Feedback>Promotion in progress>Promotion in validation>Promotion validated";
		System.out.println(text.split(">"));		
	}
	
	public static void runRegexp() {
		System.out.println("Regexp Test");
		String pattern = "^.*(\\.png|\\.css|\\.jpg|\\.ico|\\.gif|\\.js)$|^.*(\\.png|\\.css|\\.jpg|\\.ico|\\.gif|\\.js)\\?.*$";
		String header = "^image\\/.*$|^text\\/css.*$";
		
		System.out.println("/J2EEProject/js/jquery.easyui.min.js".matches(pattern));
		System.out.println("/J2EEProject/js/jquery.easyui.min.js?showRow=true".matches(pattern));
		System.out.println("/J2EEProject/login.jsp?showRow=true".matches(pattern));
		
		System.out.println("image/webp,image/*,*/*;q=0.8".matches(header));
		System.out.println("text/css,*/*;q=0.1".matches(header));
		
		String a = null;
	}

}
