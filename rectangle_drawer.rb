require 'win32ole'

class RectangleDrawer
  attr_reader :app, :book


  def initialize(excel_path)
    @app = WIN32OLE.new('Excel.Application')
    @book = @app.Workbooks.Open(excel_path)
  end

  def quit
    @book.close(false) if @book
    @app.quit if @app
  end

  def active_sheet
    @book.ActiveSheet
  end

  def shapes_in(sheet)
    array = []
    sheet.Shapes.each do |shape|
      array << shape
    end
    array
  end

end