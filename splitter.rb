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
end