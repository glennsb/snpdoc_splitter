#!/usr/bin/env ruby

require 'FileUtils'

class Dir
  def empty?
    Dir.glob("#{ path }/*", File::FNM_DOTMATCH) do |e|
      return false unless %w( . .. ).include?(File::basename(e))
    end
    return true
  end
  def self.empty? path
    new(path).empty?
  end
end

#
# Just a mapper between Snp & Investigators
#
class InvestigatorSnpMap
  
  def initialize
    @investigators_snps = {}
    @snps_investigators = {}
  end
  
  def add(inv,snp)
    add_investigator_to_snp(inv,snp)
    add_snp_to_investigator(snp,inv)
  end
  
  def investigators
    @investigators_snps.keys
  end
  
  def snps
    @snps_investigators.keys
  end
  
  def investigators_for_snp(snp)
    @snps_investigators[snp]
  end
  
  :private
  def add_investigator_to_snp(inv, snp)
    @snps_investigators[snp] ||= []
    @snps_investigators[snp] << inv
  end
  
  def add_snp_to_investigator(snp,inv)
    @investigators_snps[inv] ||= []
    @investigators_snps[inv] << snp
  end
end


#
# Interact with Microsoft Excel
#
class MsExcel
  require 'win32ole'
  
  def initialize
    @app = WIN32OLE.new('Excel.Application')
    # @app.Visible = true
    @fso = WIN32OLE.new('Scripting.FileSystemObject')
    @workbooks = []
    
    # class MsExcelConst
    # end
    WIN32OLE.const_load(@app, self.class)
    
  end
  
  def open(file)
    @workbooks << @app.Workbooks.Open(absolute_path(file))
    return @workbooks.size-1
  end
  
  def create(file)
    number_of_sheets = @app.SheetsInNewWorkbook
    @app.SheetsInNewWorkbook = 1
    @workbooks << @app.Workbooks.Add()
    @app.SheetsInNewWorkbook = number_of_sheets
    @workbooks[-1].SaveAs(absolute_path(file))
    return @workbooks.size-1
  end
  
  def sheet_of_workbook(sheet,wb_index)
    return @workbooks[wb_index].Worksheets(sheet)
  end
  
  def close(wb_index,save=false)
    # return unless @workbooks[wb_index]
    @workbooks[wb_index].Save if save
    @workbooks[wb_index].Close
    @workbooks[wb_index] = nil
  end
  
  def save_workbook(wb_index)
    @workbooks[wb_index].Save
  end
  
  def quit
    @app.Quit
  end

  def app
    @app
  end

  :private
  def absolute_path(file)
    @fso.GetAbsolutePathName(file)
  end
end


class SplitterApp
  def initialize(input,snp_select,output)
    @input_dir = input
    @snp_select_dir = snp_select
    @output_dir = output
    confirm_args()
  end
  
  def run
    @excel = MsExcel.new()
    load_investigators_snps()
    debug "Have #{@investigators_snps_map.investigators.size} investigators for #{@investigators_snps_map.snps.size} snps"
    
    # @investigators_snps_map.investigators.sort.each do |i|
    #   puts "'#{i}'"
    # end
    
    prep_output_dirs()
    
    copy_input_to_output()
    
    @excel.quit
  end
  
  :private
  def copy_input_to_output()
    files_in_dir("*.xlsx",@input_dir) do |input_file|
      input_wb = @excel.open(input_file)
      @investigators_snps_map.investigators.each do |inv|
        outputs = prep_output_files_for_input(File.basename(input_file,".xlsx"),input_wb,inv)
        search_input_to_copy_to_outputs(@excel.sheet_of_workbook(1,input_wb),outputs,inv)
        outputs.each do |key,values|
          @excel.close(values[:workbook],true)
        end
      end
      @excel.close(input_wb)
    end
  end

  def search_input_to_copy_to_outputs(source_sheet,outputs,inv)
    other_investigators = []
    (2..source_sheet.UsedRange.Rows.Count).each do |row_index|
      snp = source_sheet.Cells(row_index,1).Value.downcase.squeeze(" ").strip
      investigators = @investigators_snps_map.investigators_for_snp(snp).clone
      if investigators.delete(inv)
        other_investigators << investigators
        outputs[inv][:snps_added] += 1
        debug "Will add #{snp} from #{row_index} to #{outputs[inv][:snps_added]+1} for #{inv}"
        copy_row_from_sheet_to_sheet(row_index,source_sheet,@excel.sheet_of_workbook(1,outputs[inv][:workbook]),outputs[inv][:snps_added]+1)
      end
    end
    # insert new leading column of other investigators
    @excel.sheet_of_workbook(1,outputs[inv][:workbook]).Columns(1).Insert(MsExcel::XlToRight)
    @excel.sheet_of_workbook(1,outputs[inv][:workbook]).Cells(1,1).Value = "Contributors"
    other_investigators.each_with_index do |investigators,row|
      @excel.sheet_of_workbook(1,outputs[inv][:workbook]).Cells(2+row,1).Value = investigators.join(", ")
    end
    set_header_style(@excel.sheet_of_workbook(1,outputs[inv][:workbook]))
    
  end
  
  def prep_output_files_for_input(input_file,input_wb,investigator)
    input = @excel.sheet_of_workbook(1,input_wb)
    output = {}
    # @investigators_snps_map.investigators.each do |investigator|
      output[investigator] = {:snps_added => 0}
      output[investigator][:file] = File.join(@output_dir,investigator,"#{investigator}_#{input_file}.xlsx")      
      debug "Making new file #{output[investigator][:file]}"      
      output[investigator][:workbook] = @excel.create(output[investigator][:file])
      copy_row_from_sheet_to_sheet(1,@excel.sheet_of_workbook(1,input_wb),@excel.sheet_of_workbook(1,output[investigator][:workbook]),1)
      # set_header_style(@excel.sheet_of_workbook(1,output[investigator][:workbook]))
    # end
    output
  end
  
  def set_header_style(sheet)
    # sheet.Rows(1).Font.Bold = true
    sheet.AutoFilter()
    sheet.Rows(1).AutoFit
    (1..ws.UsedRange.Columns.Count).each do |col_index|
      sheet.Columns(col_index).AutoFit
    end
  end
  
  def copy_row_from_sheet_to_sheet(row,source_sheet,target_sheet,target_row=nil)
    target_row ||= row
    source_sheet.Rows(row).Copy
    target_sheet.Select
    target_sheet.Rows(target_row).Select
    target_sheet.Paste
  end
  
  def prep_output_dirs()
    FileUtils.mkdir(@output_dir) unless File.exists?(@output_dir)
    @investigators_snps_map.investigators.each do |i|
      FileUtils.mkdir(File.join(@output_dir,i))
    end
  end
  
  def load_investigators_snps
    @investigators_snps_map = InvestigatorSnpMap.new()
    files_in_dir("*.xls",@snp_select_dir) do |file|
      debug("Loading SNP mappings from #{file}")
      wb = @excel.open(file)
      load_snps_from_investigator_wb(wb)
      @excel.close(wb)
    end
  end
  
  def load_snps_from_investigator_wb(wb)
    each_investigators_snps_from_wb(wb) do |investigators,snp|
      investigators.split(/,/).each do |investigator|
        @investigators_snps_map.add(investigator.downcase.squeeze(" ").strip,snp.downcase.squeeze(" ").strip)
      end
    end
  end
  
  def each_investigators_snps_from_wb(wb,&block)
    ws = @excel.sheet_of_workbook(1,wb)
    (2..ws.UsedRange.Rows.Count).each do |row_index|
      investigators = ws.Cells(row_index,1).Value
      snp = ws.Cells(row_index,2).Value
      yield investigators,snp
    end
  end
  
  def files_in_dir(pattern,dir,&block)
    Dir.glob(File.join(dir,pattern)) do |file|
      yield file
    end
  end
  
  def confirm_args
    check_input_dir(@input_dir,"input")
    check_input_dir(@snp_select_dir,"SNP selection dir")
    check_output_dir()
  end
  
  def check_input_dir(dir,name)
    raise_usage "Missing an #{name} directory" unless dir
    raise_usage "Not an #{name} directory" unless File.directory?(dir)
    raise_usage "#{name} directory not readable" unless File.readable?(dir)    
  end
  
  def check_output_dir
    raise_usage "Missing an output directory" unless @output_dir
    if File.exists?(@output_dir)
      if !File.directory?(@output_dir)
        raise_usage "File already exists as output directory"
      elsif !File.writable?(@output_dir)
        raise_usage "Output directory is not writable"
      elsif !Dir.empty?(@output_dir)
        raise_usage "Output directory must be empty"
      end
    end
  end
  
  def raise_usage(msg=nil)
    STDERR.puts msg if msg
    STDERR.puts usage
    exit 1
  end
  
  def usage()
    <<-USAGE
Usage: #{File.basename(__FILE__)} INPUT_DIRECTORY SNP_SELECTION_DIRECTORY OUTPUT_DIRECTORY
    USAGE
  end
  
  def debug(msg)
    STDERR.puts "#{Time.now.utc.strftime("%Y-%m-%dT%H:%M:%SZ")}: #{msg}"
  end
  
end

if __FILE__ == $0
  SplitterApp.new(ARGV.shift,ARGV.shift,ARGV.shift).run()
end