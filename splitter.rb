#!/usr/bin/env ruby

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
  end
end

class SplitterApp
  def initialize(args)
    parse_args(args)
    confirm_args()
  end
  
  def run
    
  end
  
  :private
  def parse_args(args)
    @input_dir = args.shift
    @output_dir = args.shift
  end
  
  def confirm_args
    check_input_dir()
    check_output_dir()
  end
  
  def check_input_dir
    raise_usage "Missing an input directory" unless @input_dir
    raise_usage "Not an input directory" unless File.directory?(@input_dir)
    raise_usage "Input directory not readable" unless File.readable?(@input_dir)    
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
Usage: #{File.basename(__FILE__)} INPUT_DIRECTORY OUTPUT_DIRECTORY
    USAGE
  end
  
end

if __FILE__ == $0
  SplitterApp.new(ARGV).run()
end