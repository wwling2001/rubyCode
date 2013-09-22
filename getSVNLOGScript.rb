#encoding:utf-8
#
#测试读取log的路径
#http://bluefish-win.googlecode.com/svn/trunk/
#http://yet-another-music-application.googlecode.com/svn/trunk

#export RUBYOPT='-KU'
#require 'iconv'

require 'win32ole'
#require 'Dir'
require 'pathname'

class SvnOper

end.


class ExcelOp
    attr_reader:exApp, :exdbook, :exsheet

    def initilize()
        @exApp = WIN32OLE.new('Excel.Application')
        #@exdbook = 
    end

    def createExcelFile(fileName)
          if File.Exist(fileName) == false then
              @exdbook = @exApp.Workbooks.add()

          end
    end
end

puts File.dirname(__FILE__)
#puts Dir.pwd
puts Pathname.new(File.dirname(__FILE__)).realpath

excel = WIN32OLE.new('Excel.Application')
excel.visible = true
excel.displayalerts = false
workbook = excel.Workbooks.Add()
worksheet = workbook.Worksheets(1)
worksheet.Range("A1:D1").value = ["North","South","East","West"];
worksheet.Range("A2:B2").value = [5.2, 10]
worksheet.Range("C2").value = 8
worksheet.Range("D2").value = 20
worksheet.Range("F3").value = "中国"

range = worksheet.Range("A1:D2")
range.select
chart = workbook.Charts.Add

workbook.SaveAs("E:\\Programming\\Script\\Ruby\\myfiletest.xls")
workbook.saved = true

workbook.close()

#excel.ActiveWorkbook.Close(0)
excel.Quit()

