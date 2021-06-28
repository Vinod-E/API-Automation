from scripts.HTML_Reports.amazon_aws_s3 import AWS


class HTMLReport:

    def __init__(self, path):
        self.outputFilePath = path
        self.file = open(self.outputFilePath, "wt")

    def html_css(self, environment, sprint, date_time, use_case, result, total, success, fail, fail_color,
                 sprint_names, qa_data, beta_data, prod_data, india_data, qa_time, beta_time, prod_time,
                 india_time, success_percentage, execution_time, download_file):

        self.amazon_s3 = AWS('{}.xls'.format(sprint), download_file)
        self.amazon_s3.file_handler()

        print('**----->> Start the Printing of HTML report')
        self.file.write("""
        <html>
        <head>
        <title>""" + use_case + """</title>
        <link rel="icon" type="image/x-icon" href="https://image.flaticon.com/icons/png/512/858/858799.png"/>
        <!-- Latest compiled and minified CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">

<!-- Optional theme -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous">

<script type="text/javascript">
const myChartData = {
    type: 'doughnut',
    data: {
        labels: ["PASS", "FAIL"],
        datasets: [{
            label: "Test",
            data: """ '[' + str(success) + ',' + str(fail) + ']' + """,
            backgroundColor: [
                'green',
                'red'
            ],
            borderColor: [
                'white',
                'white'
            ],
            borderWidth: 1
        }]
    },
    options: {
      
title: {
display: true,
text: "Last run summary "
},
      animation: {
        animateScale: true,
        animateRotate: true
},
      responsive: true,
      maintainAspectRatio: false,
        
      legend: {
        position: 'right',
        labels:{
          boxWidth: 10,
          padding: 12
        }},
}};
var barChartData = {
  labels: ['AMSIN NON EU', 'AMSIN EU', 'AMS NON EU', 'AMS EU'],
  datasets: [
    {
      label: """ + '"' + sprint_names[0] + '"' + """,
      backgroundColor: "#9caad6",
      borderColor: "#9caad6",
      borderWidth: 1,
      data: """ '[' + str(qa_data[0]) + ',' + str(beta_data[0]) + ',' + str(prod_data[0]) + ',' +
                        str(india_data[0]) + ']' + """
      
    },
    {
      label: """ + '"' + sprint_names[1] + '"' + """,
      backgroundColor: "#d88589",
      borderColor: "#d88589",
      borderWidth: 1,
      data: """ '[' + str(qa_data[1]) + ',' + str(beta_data[1]) + ',' + str(prod_data[1]) + ',' +
                        str(india_data[1]) + ']' + """
    },
    {
      label: """ + '"' + sprint_names[2] + '"' + """,
      backgroundColor: "#98e1d2",
      borderColor: "#98e1d2",
      borderWidth: 1,
      data: """ '[' + str(qa_data[2]) + ',' + str(beta_data[2]) + ',' + str(prod_data[2]) + ',' +
                        str(india_data[2]) + ']' + """
    },
    {
     label: """ + '"' + sprint_names[3] + '"' + """,
      backgroundColor: "#b9de94",
      borderColor: "#b9de94",
      borderWidth: 1,
      data: """ '[' + str(qa_data[3]) + ',' + str(beta_data[3]) + ',' + str(prod_data[3]) + ',' +
                        str(india_data[3]) + ']' + """
    },
    {
     label: """ + '"' + sprint_names[4] + '"' + """,
      backgroundColor: "pink",
      borderColor: "pink",
      borderWidth: 1,
      data: """ '[' + str(qa_data[4]) + ',' + str(beta_data[4]) + ',' + str(prod_data[4]) + ',' +
                        str(india_data[4]) + ']' + """
    }
  ]
};
var timer_barChartData = {
  labels: ['AMSIN', 'BETA', 'AMS', 'INDIA'],
  datasets: [
    {
      label: """ + '"' + sprint_names[0] + '"' + """,
      backgroundColor: "#9caad6",
      borderColor: "#9caad6",
      borderWidth: 1,
      data: """ '[' + str(qa_time[0]) + ',' + str(beta_time[0]) + ',' + str(prod_time[0]) + ',' +
                        str(india_time[0]) + ']' + """
    },
    {
      label: """ + '"' + sprint_names[1] + '"' + """,
      backgroundColor: "#d88589",
      borderColor: "#d88589",
      borderWidth: 1,
      data: """ '[' + str(qa_time[1]) + ',' + str(beta_time[1]) + ',' + str(prod_time[1]) + ',' +
                        str(india_time[1]) + ']' + """
    },
    {
      label: """ + '"' + sprint_names[2] + '"' + """,
      backgroundColor: "#98e1d2",
      borderColor: "#98e1d2",
      borderWidth: 1,
      data: """ '[' + str(qa_time[2]) + ',' + str(beta_time[2]) + ',' + str(prod_time[2]) + ',' +
                        str(india_time[2]) + ']' + """
    },
    {
     label: """ + '"' + sprint_names[3] + '"' + """,
      backgroundColor: "#b9de94",
      borderColor: "#b9de94",
      borderWidth: 1,
      data: """ '[' + str(qa_time[3]) + ',' + str(beta_time[3]) + ',' + str(prod_time[3]) + ',' +
                        str(india_time[3]) + ']' + """
    },
    {
     label: """ + '"' + sprint_names[4] + '"' + """,
      backgroundColor: "pink",
      borderColor: "pink",
      borderWidth: 1,
      data: """ '[' + str(qa_time[4]) + ',' + str(beta_time[4]) + ',' + str(prod_time[4]) + ',' +
                        str(india_time[4]) + ']' + """
    }
  ]
};
var chartOptions = {
  responsive: true,
  legend: {
    position: "bottom"
  },
  egend: {
      position: 'bottom',
      labels: {
         usePointStyle: true
      }
   },
  title: {
    display: true,
    text: "Last 5 Sprints (Use Cases) Report"
  },
  scales: {
    yAxes: [{
    scaleLabel: {
        display: true,
        labelString: 'Use Cases'
      },
      ticks: {
        beginAtZero: true
      }
    }]
  }
}
var chart_time_Options = {
  responsive: true,
  legend: {
    position: "bottom"
  },
  title: {
    display: true,
    text: "Last 5 Sprints (Time Taken) Report"
  },
  scales: {
    yAxes: [{
    scaleLabel: {
        display: true,
        labelString: 'Time (min)'
      },
      ticks: {
        beginAtZero: true
      }
    }]
  }
}
window.onload = function() {
  var ctx = document.getElementById("canvas").getContext("2d");
  var ctx_time = document.getElementById("canvas_time").getContext("2d");
  var myChart123 = document.getElementById("myChart123").getContext("2d");
  window.myBar = new Chart(ctx, {
    type: "bar",
    data: barChartData,
    options: chartOptions
  });
  window.myBar = new Chart(ctx_time, {
    type: "bar",
    data: timer_barChartData,
    options: chart_time_Options
  });
  window.myBar = new Chart(myChart123, {
    type: myChartData.type,
    data: myChartData.data,
    options: myChartData.options
  });
};

</script> 
<style>

    body {
        font-family: Inter, Segoe UI, Roboto, Arial, verdana, geneva, sans-serif;
        background: #f5f2f2;

    }
    .tableStyle{
    width: 100%;
    max-width: 100%;
    margin-bottom: 20px;
    font-family: Inter,Segoe UI,Roboto,Arial,verdana,geneva,sans-serif;
    font-size: 12px;
    font-style: normal;
    font-weight: 600;
    line-height: 16px;
    text-align: left;
    color: #6b6b6b;
    }
    .alink {
    cursor: pointer;
    color: #0265d2;
    text-decoration: none;
}
.tbltd{
word-break:break-all;word-wrap:break-word;overflow:hidden;font-family:Inter,Segoe UI,Roboto,Arial,verdana,geneva,sans-serif;font-size:12px;font-style:normal;font-weight:400;line-height:16px;text-align:left;color:#000
}
.custBtn{
background: #1f8ae7 !important;
border-color: #1f8ae7 !important;
font-weight: 600;
}
   .subHeader{
   margin-top: 1rem;
   font-size:20px;
   font-weight:700;
   line-height:24px;
   text-align:left;
   letter-spacing:-.56px;
   color: #6b6b6b;
   }
   .footer{
   font-size: 11px;
   font-weight: 400;
   line-height: 18px;
   text-align: left;
   color: #a6a6a6;
   margin-top: 1rem;
   margin-bottom: 3rem;
   }
   .summHeader {
        font-size: 16px;
        font-weight: 700;
        line-height: 20px;
        letter-spacing: -.12px;
        text-align: left
    }
    .Pass{
    color: green;
    font-size: 17px !important;
    font-weight: bold !important;
}
.Fail{
    color: red;
    font-size: 17px !important;
    font-weight: bold !important;
}
.summaryPass{
color: green;
font-weight: bold !important;
}
.summaryFail{
color: red;
font-weight: bold !important;
}
</style>
</head>
<body>
<div class="container" style="background: #ffff !important;border: 1px solid #e2e2e2;margin-top: 1rem; margin-bottom: 1rem;">
<div class="row" style="    padding: 20px;">
<a>
<img width="142"
src="https://hirepro.in/wp-content/uploads/2020/08/hirepro-new-logo-dark-slim.png"
style="border:0;display:block;outline:0;text-decoration:none;height:auto">
</a>
    <div class="row" style="margin-top: 3rem">
<div class="col-xs-12 col-sm-12 col-lg-6 subHeader"> Automated Test Reporting
</div>
<div class="col-xs-12 col-sm-12 col-lg-6">
<div class="btn-toolbar" style="float: right;">

<a href=""  
type="button" id="btnSubmit" title="HTML Web Report" class="btn btn-primary btn-sm custBtn"><img  
src="https://image.flaticon.com/icons/png/512/2353/2353373.png" width="25" height="25"/> View Run Results</a>


<a href=""" + self.amazon_s3.one_day_link + """ target="_blank" 
type="button" id="btnCancel" title="Download Excel Report" class="btn btn-primary btn-sm custBtn"><img  
src="https://image.flaticon.com/icons/png/512/1053/1053166.png" width="25" height="25"/> Excel Download</a>


<a title="sprint wise automation reports" 
href="https://drive.google.com/drive/u/1/folders/186nL7DWI_ZoMklgcwIUykC4tSQuECtGH" target="_blank" type="button" 
id="btnCancel" class="btn btn-primary btn-sm custBtn"><img  
src="https://image.flaticon.com/icons/png/512/2965/2965323.png" width="25" height="25"/> Google Drive</a>
				</div>
        	</div>
            </div>
		</div>
      <div class="row row-offcanvas row-offcanvas-right">

        <div class="col-xs-12 col-sm-12">
          <p class="pull-right visible-xs">
            <button type="button" class="btn btn-primary btn-xs" data-toggle="offcanvas">Toggle nav</button>
          </p>
         
          <div class="row">
            <div class="col-xs-12 col-lg-6 ">
             <table class="table table-sm tableStyle" >
			  
			  <tbody>
			    <tr>			     
			      <td>Monitor</td>
			      <td>	
			      	<div class="tbltd">
                        <a class="alink">Vinod Eraganaboina</a>
                    </div>
				  </td>
			    </tr>
			    <tr>
			      
			      <td>Workspace</td>
					<td>	
			      	<div class="tbltd">
                        <a class="alink">UI Automation</a>
                    </div>
				  </td>
			    </tr>
			    <tr>
			      
			      <td>Environment</td>
			      <td>	
	
					<div class="tbltd">
                        <a class="alink">""" + environment + """</a>
                    </div>
					</td>
			      
			    </tr>
			     <tr>
			      
			      <td>Sprint</td>
			      <td>	
					<div class="tbltd">
                        <a class="alink">""" + sprint + """</a>
                    </div>
                	</td>
			      
			    </tr> 
			    <tr style="border-bottom: 1px solid #ddd;"> 		      
			      <td>Run Date&Time</td>
			      <td>	
					<div class="tbltd">
                        <a class="alink">""" + str(date_time) + """</a>
                    </div>
                </td>
			    </tr>
			  </tbody>
			</table>
            </div>
            <div class="col-xs-12 col-lg-6">
              <table class="table table-sm tableStyle" >
			  
			  <tbody>
			    <tr>			     
			      <td>Collection</td>
			      <td>	
			      	<div class="tbltd">
                        <a class="alink">""" + use_case + """</a>
                    </div>
				  </td>
			    </tr>
			    <tr>
			      
			      <td>Execution Time</td>
					<td>	
			      	<div class="tbltd">
                        <a class="alink">""" + str(execution_time) + """ min</a>
                    </div>
				  </td>
			    </tr>
			    <tr>
			      <td>Total Cases</td>
			      <td>	
	
					<div class="tbltd">
                        <a class="alink">""" + str(total) + """</a>
                    </div>
					</td>
			    </tr>
			    <tr>
			      <td>Success %</td>
			      <td>	

					<div class="tbltd">
						<div style="font-size:12px;font-style:normal;line-height:16px;
						text-align:left" class=""" + result + """>""" + str(success_percentage) + """  %</div>
					</td>
			    </tr>
			    <tr style="border-bottom: 1px solid #ddd;"> 
			      <td>Result</td>
			      <td>	
					<div class="tbltd">
                        <div style="font-size:12px;font-style:normal;line-height:16px;
                        text-align:left" class=""" + result + """>""" + result + """</div>
                    </div>
                </td>
			    </tr>
			  </tbody>
			</table>
            </div>
          </div>
        </div>
      </div>
     
       <div class="row row-offcanvas row-offcanvas-right">
       <div style="margin-bottom: 0rem;text-align: left;margin-top: 3rem;margin-left: 2rem;">
         <div class="summHeader" style="color: #6b6b6b;">Historical Summary</div>
     	</div>
	        <div class="col-xs-12 col-sm-12 col-lg-6">
			<div style="padding:32px 48px 0px 40px;">
				<div id="container" style="margin-top:2rem;height: 300px; width: 100%;">
					<canvas id="canvas"></canvas>
					<!-- <div style="font-size:12px;font-family: 'Helvetica Neue', 'Helvetica', 'Arial', sans-serif;
					font-weight: 600;text-align: center;"> <p> Chart legends are clickable to view the specific items </p> </div> -->
				</div>
			</div>
	        </div>
	        <div class="col-xs-12 col-sm-12 col-lg-6">
	        	<div style="padding:32px 48px 0px 40px;">
                <div id="container" style="margin-top:2rem;height: 300px; width: 100%;">
                  <canvas id="canvas_time"></canvas>
                  <!-- <div style="font-size:12px;font-family: 'Helvetica Neue', 'Helvetica', 'Arial', sans-serif;
    font-weight: 600;text-align: center;"> <p> Chart legends are clickable to view the specific items </p> </div> -->
                </div>
            </div>
    	</div>
  </div>
   <div class="row row-offcanvas row-offcanvas-right">
   	<div class="col-xs-12 col-sm-12 col-lg-6">
   		<div style="margin-bottom: 3rem;text-align: left;margin-top: 0rem;">
         <div class="summHeader" style="color: #6b6b6b;"> Last run summary</div>
     	</div>
   		<table class="table table-sm tableStyle" >
			 <tbody>
                <tr>
                    <td class="emptyCell">
                        &nbsp;</td>
                    <td class="summaryHeads"> Passed</td>
                    <td class="summaryHeads">Failed</td>
                    <td class="summaryHeads">Total</td>
                </tr>
                <tr>
                <td class="summaryHeads" style="border-bottom:1px solid #ededed;">Tests</td>
                <td class="summaryRow summaryPass" style="border-bottom:1px solid #ededed;">""" + str(success) + """</td>
                <td style="border-bottom:1px solid #ededed;" class="summaryRow """ + fail_color + """ ">""" + str(fail) + """</td>
                <td style="border-bottom:1px solid #ededed;" class="summaryRow" style="font-weight:bold;color: black;">""" + str(total) + """</td>
                </tr>
            </tbody>
			</table>
   	</div>
   	<div class="col-xs-12 col-sm-12 col-lg-6"> 
   		<div style="padding:32px 48px 0px 40px;">
                <div id="container" style="height: 150px; width: 100%;">
                  <canvas id="myChart123"></canvas>
                  <!-- <div style="font-size:12px;font-family: 'Helvetica Neue', 'Helvetica', 'Arial', sans-serif;
    font-weight: 600;text-align: center;"> <p> Chart legends are clickable to view the specific items </p> </div> -->
                </div>
            </div>
   	</div>
   </div>
<div style="padding:2rem 0 15px 0;">
                                <p style="border-top:solid 2px #ededed;font-size:1px;margin:0 auto;width:100%">
                                </p>
                            </div>
<footer class="footer">
      <div class="container">
        <span class="text-muted">© 2021 HirePro . All rights reserved.
<span>Plot No. 53, Kariyammana Agrahara Road,Devarabisana Halli, Bengaluru – 560 103</span>
        </span>
      </div>
    </footer>
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.3.0/Chart.bundle.js"></script>
</body>
</html>
        """)
