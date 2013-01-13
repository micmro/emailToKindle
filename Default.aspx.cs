using System;
using System.Collections.Generic;


public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //kindle email address
        string kindleEmailAddress = "yourKinldeName@kindle.com";
        List<OpenPop.Mime.Message> messages = SendEmailAsAttachment.FetchMessagesAndSendAsAttachment(kindleEmailAddress);

        //to add http://blog.stackoverflow.com/2008/07/easy-background-tasks-in-aspnet/
        // or http://quartznet.sourceforge.net/


        //render email body
        emails.DataSource = messages;
        emails.DataBind();
    }

   
}
