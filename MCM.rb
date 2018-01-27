require 'creek'

workbook = Creek::Book.new 'MCM3.xlsx'
worksheets = workbook.sheets


class Topic
	attr_accessor :name
	attr_accessor :pointsDescription
	def initialize(name, points, description)
		@name = name
		@pointsDescription = [[points, description, checkGood(points, description)]]
	end

	def addPointsDescription(points, description)
		@pointsDescription.push([points, description, checkGood(points, description)])
	end

	def checkGood(points, description)
		if ((@name == "Visible tool" && points.to_i == 1) ||(@name != "Visible tool" && points.to_i > 1)) #pass
			return true

		else #fail
			return false
		end
	end

	def reversePoints()
		@pointsDescription = @pointsDescription.reverse
	end
end

topics = []
current_topic = nil
worksheets.each do |worksheet|
  worksheet.rows.each do |row|
    row_cells = row.values
    #new topic
    if (row_cells[0] != nil) 
    	if (current_topic != nil)
    		current_topic.reversePoints()
    		topics.push(current_topic)
    	end
    	current_topic = Topic.new(row_cells[0], row_cells[2][0], row_cells[4].gsub("'", "\\\\'"))
    else
    	current_topic.addPointsDescription(row_cells[2][0], row_cells[4])
    end
    # do something with row_cells
  end
end

html = "<!DOCTYPE html>
<html>
	<head>
	    <script src='script.js'></script>
	</head>

<form id = 'formL1' style='height: 1000px;'>
<div style='margin-left: 10%; float:left;'>\n"
javascript = "var goodParagraph = '';"
javascript2 = 'function feedback()
{
	var form = document.getElementById("formL1");
	var goodParagraph = "";
	var badParagraph = "";'
for i in 0...topics.length
	if (i == topics.length/2)
		html += "</div>\n<div style='margin-left: 30%; float: left; margin-right: 10%'>\n"	
	end
	html += "<div id = '#{topics[i].name}'>\n<h3> #{topics[i].name} </h3>\n"
	javascript2 += "var lookup = '#{topics[i].name.gsub(/\s+/, "")}' + form.#{topics[i].name.gsub(/\s+/, "")}.value;\nif (window[lookup][1])\n{\ngoodParagraph += '#{topics[i].name}' + ': ' + window[lookup][0] + '<br><br>';\n}\nelse\n{\nbadParagraph += '#{topics[i].name}' + ': ' + window[lookup][0] + '<br><br>';\n}\n"
	for j in 0...topics[i].pointsDescription.length
		javascript += "var #{topics[i].name.gsub(/\s+/, "")}#{topics[i].pointsDescription[j][0]} = [\"#{topics[i].pointsDescription[j][1]}\", #{topics[i].pointsDescription[j][2]}]\n"
		html += "<input type='radio' name='#{topics[i].name.gsub(/\s+/, "")}' value='#{topics[i].pointsDescription[j][0]}'>#{topics[i].pointsDescription[j][0]}\n"
	end
	html += "</div>\n"
end

html += "</div>\n<button id='feedbackButton' type='button' onclick='feedback()'>Feedback</button>\n</form>\n<div style='margin-top: 1000px;'>\n<h3> Feedback </h3>\n<div id = 'goodParagraph' style='float: left; width:50%'>\nGood Paragraph\n</div>\n<div id = 'badParagraph' style='float:left; width: 50%;'>\nImprovement Paragraph\n</div>\n</div>\n</html>"

puts html
puts javascript + javascript2
javascript2 += 'var good = document.getElementById("goodParagraph");
	var bad = document.getElementById("badParagraph");
	good.innerHTML=goodParagraph;
	bad.innerHTML=badParagraph;
	window.location="#goodParagraph"}'

File.write('index.html', html)
File.write('script.js', javascript + javascript2)



