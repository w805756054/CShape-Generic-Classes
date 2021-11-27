// See https://aka.ms/new-console-template for more information


using IanWongHelpers;
using System.ComponentModel;
using System.Net.Mail;


var host = "smtp.sendgrid.net";
var username = "apikey";
var password = "your-password";
var port = 587;
var enableSSL = true;

var smtpHelper = new SmtpHelper(host, port, enableSSL, username, password);
var subject = "Subject for test";
var body = "<div style=\"font-size:20px\">Body for test...</div>";
var from = "xxx@qq.hk";
var to = "xxx@qq.com";
var fileAttachment = new System.Net.Mail.Attachment("d:/textfile.txt");
smtpHelper.SendAsync(subject, body, true, from, to, new List<Attachment>() { fileAttachment });

Console.ReadLine();

