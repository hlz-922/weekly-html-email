import pandas as pd

date = input('Sendout date [YYYY-MM-DD]: ')
html_folder_path = ''
excel_folder_path = ''
html = open(html_folder_path + date + '.html', 'w+')
df = pd.read_excel(excel_folder_path + 'Weekly Bulletin.xlsx')
df = df.T.reset_index(drop=True).T
# print(df)

# add CSS details
formatting = '''
<!doctype html>
<html>
    <head>
    <meta charset="UTF-8">
    <!-- utf-8 works for most cases -->
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Forcing initial-scale shouldn't be necessary -->
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <!-- Use the latest (edge) version of IE rendering engine -->
    <title>  </title>
    <!-- The title tag shows in email notifications, like Android 4.4. -->
    <!-- Please use an inliner tool to convert all CSS to inline as inpage or external CSS is removed by email clients -->
    <!-- important in CSS is used to prevent the styles of currently inline CSS from overriding the ones mentioned in media queries when corresponding screen sizes are encountered -->

    <!-- CSS Reset -->
    <style type="text/css">
/* What it does: Remove spaces around the email design added by some email clients. */
      /* Beware: It can remove the padding / margin and add a background color to the compose a reply window. */
html,  body {
	margin: 0 !important;
	padding: 0 !important;
	height: 100% !important;
	width: 100% !important;
}
/* What it does: Stops email clients resizing small text. */
* {
	-ms-text-size-adjust: 100%;
	-webkit-text-size-adjust: 100%;
}
/* What it does: Forces Outlook.com to display emails full width. */
.ExternalClass {
	width: 100%;
}
/* What is does: Centers email on Android 4.4 */
div[style*="margin: 16px 0"] {
	margin: 0 !important;
}
/* What it does: Stops Outlook from adding extra spacing to tables. */
table,  td {
	mso-table-lspace: 0pt !important;
	mso-table-rspace: 0pt !important;
}
/* What it does: Fixes webkit padding issue. Fix for Yahoo mail table alignment bug. Applies table-layout to the first 2 tables then removes for anything nested deeper. */
table {
	border-spacing: 0 !important;
	border-collapse: collapse !important;
	table-layout: fixed !important;
	margin: 0 auto !important;
}
table table table {
	table-layout: auto;
}
/* What it does: Uses a better rendering method when resizing images in IE. */
img {
	-ms-interpolation-mode: bicubic;
}
/* What it does: Overrides styles added when Yahoo's auto-senses a link. */
.yshortcuts a {
	border-bottom: none !important;
}
/* What it does: Another work-around for iOS meddling in triggered links. */
a[x-apple-data-detectors] {
	color: inherit !important;
}
</style>

    <!-- Progressive Enhancements -->
    <style type="text/css">

        /* What it does: Hover styles for buttons */
        .button-td,
        .button-a {
            transition: all 100ms ease-in;
        }
        .button-td:hover,
        .button-a:hover {
            background: #555555 !important;
            border-color: #555555 !important;
        }
hr {
    border-right : 0;
    border-left: 0;
}
        /* Media Queries */
        @media screen and (max-width: 600px) {

            .email-container {
                width: 100% !important;
            }

            /* What it does: Forces elements to resize to the full width of their container. Useful for resizing images beyond their max-width. */
            .fluid,
            .fluid-centered {
                max-width: 100% !important;
                height: auto !important;
                margin-left: auto !important;
                margin-right: auto !important;
            }
            /* And center justify these ones. */
            .fluid-centered {
                margin-left: auto !important;
                margin-right: auto !important;
            }

            /* What it does: Forces table cells into full-width rows. */
            .stack-column,
            .stack-column-center {
                display: block !important;
                width: 100% !important;
                max-width: 100% !important;
                direction: ltr !important;
            }
            /* And center justify these ones. */
            .stack-column-center {
                text-align: center !important;
            }

            /* What it does: Generic utility class for centering. Useful for images, buttons, and nested tables. */
            .center-on-narrow {
                text-align: center !important;
                display: block !important;
                margin-left: auto !important;
                margin-right: auto !important;
                float: none !important;
            }
            table.center-on-narrow {
                display: inline-block !important;
            }

        }

    </style>
    </head>
    <body bgcolor="#e0e0e0" width="100%" style="margin: 0;" yahoo="yahoo">
    <table bgcolor="#e0e0e0" cellpadding="0" cellspacing="0" border="0" height="100%" width="100%" style="border-collapse:collapse;">
      <tr>
        <td><center style="width: 100%;">

            <!-- Email Header : BEGIN -->
            <table align="center" width="600" class="email-container">
            <tr>
                <td style="padding:20px; text-align: center"><img src="images/logo.png" align="center" width="200" height="50" alt="alt_text" border="0"></td>
              </tr>
          </table>
            <!-- Email Header : END -->

            <!-- Email Body : BEGIN -->
            <table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#ffffff" width="600" class="email-container">

            <!-- Hero Image, Flush : BEGIN -->
            <tr>
                <td class="full-width-image" style="padding: 30px"><img src="images/dining-hall-crayon.png" width="600" alt="alt_text" border="0" align="center" style="width: 100%; max-width: 600px; height: auto;"></td>
              </tr>
            <!-- Hero Image, Flush : END -->
'''
html.write(formatting)





# find my intro from excel file
intro_info = df.iloc[1:len(df),0:2]
intro_info = intro_info.dropna()
intro_info.columns = ['message','order']
intro_info.order = pd.to_numeric(intro_info.order)
intro_info = intro_info.sort_values(by=['order'])
intro_info = intro_info.reset_index(drop=True)

# write my intro into html
intro = '''
<!-- 1 Column Text : BEGIN -->
            <tr>
              <td style="padding: 10px 40px 40px 40px; text-align: left; font-family: sans-serif; font-size: 15px; mso-height-rule: exactly; line-height: 20px; color: #555555;">
					<p><strong>Dear BAs, </strong></p>'''

for i in range(0,len(intro_info)):
    temp = '''<p>'''+intro_info.iloc[i,0]+'''</p>'''
    intro = intro + temp

intro= intro+'''
<p>You could subscribe to our event calendar using this <a href="https://outlook.office365.com/owa/calendar/69ae0c0db1b7490aaa8293870f6b80fd@trin.cam.ac.uk/5b655da2b70348218a97f10c96431d0410329026302815681283/calendar.ics">link</a> or view it on the <a href="https://outlook.office365.com/calendar/published/69ae0c0db1b7490aaa8293870f6b80fd@trin.cam.ac.uk/5b655da2b70348218a97f10c96431d0410329026302815681283/calendar.html">web</a>. Please always refer to the bulletin because the calendar might not be up-to-date; and I'd recommend choosing a reasonable auto-refresh rate, for instance, 1 day. </p>
<p>Send us feedbacks via the <a href="https://forms.office.com/Pages/ResponsePage.aspx?id=4Kd2NkFfFk2b51fqbRWgK6CmBfFmLLtAqzZAd9Kkg2xUMzlDVkRRWkpaV0FCUEJZQjA3SkdDT1dPTS4u">contact form</a>, or email specific committee members by clicking <a href="http://basociety.net/committee/">here</a>. :) </p>
<p>If you wish to make formal or informal complains to college about any misconduct, please consult <a href="https://students.trin.cam.ac.uk/respect-dignity-and-inclusion/">this website</a> for more details. Also, feel free to reach out to any committee member whom you feel comfortable to talk with - we're all here to help.</p>
<p><strong>All the best,<br>
Lillian</strong></p></td>
<br>
</tr>
'''
html.write(intro)






# find special announcement from excel file
sa = df.iloc[1:len(df),4:7]
sa = sa.dropna()
sa.columns = ['title','message','order']
sa.order = pd.to_numeric(sa.order)
sa = sa.sort_values(by=['order'])
sa = sa.reset_index(drop=True)

# add special announcement
for i in range(0,len(sa)):
    special = '''
    <!-- special announcement : BEGIN -->
            <tr>
                <td bgcolor="#C6878F" valign="middle" style="text-align: center;">
                    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td valign="middle" style="text-align: center; padding: 5px; font-family: sans-serif; font-size: 15px; mso-height-rule: exactly; line-height: 20px; color: #ffffff;"> <h3><strong> '''+sa.iloc[i,0]+''' </strong></h3></td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
              <td style="padding: 40px; text-align: left; font-family: sans-serif; font-size: 15px; mso-height-rule: exactly; line-height: 20px; color: #555555;">
				<p>'''+sa.iloc[i,1]+'''</p>
			  </td>
                <br>
			</tr>
			<!-- special announcement : END -->	'''
    html.write(special)






# find events from excel file
event = df.iloc[1:len(df),9:21]
event = event.dropna(how='all')
event.columns = ['email','nah','nah2','order','image','type','title','date','contact','signup','join','description']
event.order = pd.to_numeric(event.order)
event = event.sort_values(by=['order'])
event = event.dropna(subset=['order'])
event = event.reset_index(drop=True)
# print(event.iloc[:,3])

# add summary
summary = '''
            <tr>
                <td bgcolor="#C6878F" valign="middle" style="text-align: center;">
                    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td valign="middle" style="text-align: left; padding: 40px; font-family: sans-serif; font-size: 15px; mso-height-rule: exactly; line-height: 20px; color: #ffffff;">
							<p><strong> Happening soon </strong></p>
							<ul>
'''
for i in range(0,len(event)):
    if event.iloc[i,5] == 'h':
        summary = summary+'''<li>'''+event.iloc[i,6]+''' --- '''+event.iloc[i,7]+'''</li>'''
summary = summary+'''</ul><br>
							<strong><p> Recurring events </p></strong>
							<ul>'''
for i in range(0,len(event)):
    if event.iloc[i,5] == 'r':
        summary = summary+'''<li>'''+event.iloc[i,6]+''' --- '''+event.iloc[i,7]+'''</li>'''
summary = summary+'''</ul>
						</td>
					</tr>
					</table>
				</td>
			</tr>'''
html.write(summary)






# add events
eb='''
<tr>
                <td align="center" valign="top" style="padding: 40px 0 0 0;">
				<table cellspacing="0" cellpadding="0" border="0" width="100%">
'''
signup = pd.isna(event['signup'])
join = pd.isna(event['join'])
for i in range(0,len(event)):
    eb = eb+'''
					<tr>
                    <td class="stack-column-center">
						<table cellspacing="0" cellpadding="0" border="0">
                        	<tr>
                        		<td style="padding: 10px 40px; text-align: center;"><img src="images/events/'''+event.iloc[i,4]+'''.png" width="120" height="120" alt="alt_text" border="0" class="fluid"></td>
                        		<td style="font-family: sans-serif; font-size: 15px; mso-height-rule: exactly; line-height: 20px; color: #555555; padding: 10px 40px; text-align: left; width:100%" class="center-on-narrow">
								<br><strong>'''+event.iloc[i,6]+'''</strong><br>
								Date: '''+event.iloc[i,7]+'''<br>
								Contact: <a href="mailto:'''+event.iloc[i,0]+'''">'''+event.iloc[i,8]+'''</a><br>'''

    if signup[i] == False:
        eb = eb+'''
								Click <a href="'''+event.iloc[i,9]+'''">here</a> to sign up <br>'''
    if join[i] == False:
        eb = eb+'''
								Click <a href="'''+event.iloc[i,10]+'''">here</a> to join online'''
    eb = eb+'''</td>
                      		</tr>
                      	</table></td>
                  	</tr>
					<tr>
					<td style="font-family: sans-serif; font-size: 15px; mso-height-rule: exactly; line-height: 20px; color: #555555; padding: 30px; text-align: left" class="center-on-narrow">
						<p>'''+event.iloc[i,11]+'''</p>
					  <br><br><hr></td>
					</tr>'''

html.write(eb)







# find extra messages from excel file
em = df.iloc[1:len(df),25:29]
em = em.dropna()
em.columns = ['title','contact','message','order']
em.order = pd.to_numeric(em.order)
em = em.sort_values(by=['order'])
em = em.reset_index(drop=True)

# add extra messages
if len(em)>0:
    extra = '''<!-- special announcement : BEGIN -->
            <tr>
                <td bgcolor="#C6878F" valign="middle" style="text-align: center;">
                    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td valign="middle" style="text-align: center; padding: 5px; font-family: sans-serif; font-size: 15px; mso-height-rule: exactly; line-height: 20px; color: #ffffff;"> <h3><strong> Circulated messages from outside BA society </strong></h3></td>
					</tr>
					</table>
				</td>
			</tr>'''
    for i in range(0,len(em)):
        extra = extra+'''
			<tr>
              <td style="padding: 40px; text-align: left; font-family: sans-serif; font-size: 15px; mso-height-rule: exactly; line-height: 20px; color: #555555;">
				<h3>'''+em.iloc[i,0]+'''</h3>
				  <p><strong>by '''+em.iloc[i,1]+'''</strong></p>
				  <p>'''+em.iloc[i,2]+'''</p><br><hr><br>'''
    extra = extra+'''</td>
			</tr>
			<!-- special announcement : END -->	'''
    html.write(extra)


2020




# add final bits
end = '''
</table></td>

</tr></table>
            <!-- Email Body : END -->

            <!-- Email Footer : BEGIN -->
            <table align="center" width="600" class="email-container">

            <tr>
				<td align="center" valign="top" style="padding: 10px;">
				<table cellspacing="0" cellpadding="0" border="0" width="100%">
                <td width="100%" class="stack-column-center">
					<table cellspacing="0" cellpadding="0" border="0">
                        <tr>
                        <td style="padding: 20px;"><a href="https://www.facebook.com/groups/TrinityBASociety"><img src="images/icons/fb.png" width="50" height="50" alt="alt_text" border="0" class="fluid"></a></td>
						<td style="padding: 20px;"><a href="https://www.instagram.com/trin_ba_society/"><img src="images/icons/ins-logo-brown.png" width="80" height="80" alt="alt_text" border="0" class="fluid"></a></td>
                        <td style="padding: 20px;"><a href="http://basociety.net/"><img src="images/icons/weblink.png" width="50" height="50" alt="alt_text" border="0" class="fluid"></a></td>
                 		</tr>
					</table>
                  	</td>
				</table>
			    </td>
             </tr>
			 <tr>
				<td style="padding: 30px 10px;width: 100%;font-size: 12px; font-family: sans-serif; mso-height-rule: exactly; line-height:18px; text-align: center; color: #888888;">
				<strong>Know your committee!</strong><br>This <a href="http://basociety.net/committee/">link</a> will direct you to the webpage that contains the email address to our committee members.
				<br><br>
				<strong>Events cancellation policy</strong> <br>After the sign-up has closed, we will notify you as soon as possible if you got a place. Then, you will have one full day (24 hours) to decide if you (and your guest) want to accept the place or if you want to drop out free of charge. If you (or your guest) drop-out later with at least 24 hours notice before the event and there is no one on the waiting list, you will still get charged the subsidised price. If you (or your guest) do not show up to the event without giving any notice, or you give notice and there is less than 24 hours to the event start time, you will be charged the full (not subsidised) price.
				<br><br>
                <strong>BA society, Trinity college, University of Cambridge</strong><br><br>
                <span class="mobile-link--footer"></span>HTML email developed by computing officer (H.L.Z.) June 2020<br>
                <br><unsubscribe style="color:#888888; text-decoration:underline;"><a href="https://lists.cam.ac.uk/">Click here to manage your email lists</a></unsubscribe>
                </td>
              </tr>
          </table>
            <!-- Email Footer : END -->

          </center></td>
      </tr>
    </table>
</body>
</html>
'''
html.write(end)


html.close()
