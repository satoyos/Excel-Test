require_relative '../rectangle_drawer'
require_relative '../../../lib/rspec_helper_for_windows'

describe RectangleDrawer do
  RIGHT_PATH = 'D:\src\Ruby\test\Excel-Test\test1.xls'
  WRONG_PATH = 'D:\TEST_WRONG_PATH\test1.xls'
  describe '初期化' do
    context '正しいパスでの初期化時' do
      subject{RectangleDrawer.new(RIGHT_PATH)}

      it {should_not be_nil}
      its(:app){should_not be_nil}
      its(:book){should_not be_nil}

      after do
        subject.quit
      end
    end

    context '誤ったパスでの初期化時' do
      subject{lambda{RectangleDrawer.new(WRONG_PATH)}}

      it {should raise_error}

    end
  end

  describe 'active_sheet' do
    subject{RectangleDrawer.new(RIGHT_PATH)}

    its(:active_sheet){should_not be_nil}

    after do
      subject.quit
    end
  end

  describe 'shapes in sheet' do
    before do
      @drawer = RectangleDrawer.new(RIGHT_PATH)
      @sheet = @drawer.active_sheet
    end

    it 'should be a Array of shapes' do
      @drawer.shapes_in(@sheet).should_not be_nil
      @drawer.shapes_in(@sheet).should be_instance_of(Array)
      @drawer.shapes_in(@sheet).each_with_index do |shape, idx|
        puts "text of shape[#{idx}] => [#{shape.AlternativeText}]"
      end
      1.should == 1
    end

    after do
      @drawer.quit
    end
  end

end