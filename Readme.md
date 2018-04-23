# How to export scheduler appointments in Google Calendar


<p>This example demonstrates how you can use the <strong>Google Calendar API</strong> in your scheduling application. Google provides the corresponding guidelines regarding use of this API:<br><a href="https://developers.google.com/google-apps/calendar/quickstart/dotnet">Google Calendar API</a> </p>
<p>Before using this API, make certain you have read and understand <a href="https://developers.google.com/site-policies">Google’s licensing terms</a>. Next, you’ll need to generate a corresponding JSON file with credentials to enable the <strong>Google Calendar API.</strong></p>
<p>We have a corresponding KB article which contains step-by-step description on how to generate this JSON file:<br><a href="https://www.devexpress.com/Support/Center/p/T267842">How to enable the Google Calendar API to use it in your application</a><br><br></p>
<p>After you generate this JSON file and put it in the <strong>Secret </strong>folder of this sample project, you can export appointments to a Google calendar.<br>1. Click the "<strong>Connect</strong>" button to generate a list of available calendars for your Google account and load Google events to the SchedulerControl from the selected calendar.<br>2. Add or change an appointment in the SchedulerControl. The modified appointment will be exported to the Google Calendar.</p>
<p><br><br></p>
<p><strong>P.S. To run this example's solution, include the corresponding "Google Calendar API" assemblies into the project.</strong><br><strong>For this, open the "Package Manager Console" (Tools - NuGet Package Manager) and execute the following command:<br></strong></p>
<pre class="prettyprint notranslate"><code>Install-Package Google.Apis.Calendar.v3<br>Install-Package NodaTime -Version 1.3.1 </code></pre>


<h3>Description</h3>

<p>For exporting&nbsp;appointments, we created a corresponding&nbsp;<strong>CalendarExporter</strong>&nbsp;class which creates&nbsp;<strong>an</strong>&nbsp;<strong>Event (Google.Apis.Calendar.v3.Data)</strong>&nbsp;based on an&nbsp;<strong>Appointment</strong>&nbsp;instance. The&nbsp;<strong>GoogleCalendarUtils&nbsp;</strong>class-helper&nbsp;copies properties of the scheduler appointment according to&nbsp;Google Calendar API specifics. For recurrent appointments, the pattern's EventId is checked. If there is no recurring event in the Google Calendar for this instance, a new recurring event is created. The&nbsp;<strong>VRecurrenceConverter</strong>&nbsp;class converts the appointment pattern's recurrence information in the Google recurring event's rule.&nbsp;</p>

<br/>


