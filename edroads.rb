#!/usr/bin/ruby -w
#============================================================================
#	NAME:
#		edroads.rb
#
#	DESCRIPTION:
#		Edinburgh Council road processor.
#	
#	COPYRIGHT:
#		Copyright (c) 2010-2019, refNum Software
#		All rights reserved.
#
#		Redistribution and use in source and binary forms, with or without
#		modification, are permitted provided that the following conditions
#		are met:
#		
#		1. Redistributions of source code must retain the above copyright
#		notice, this list of conditions and the following disclaimer.
#		
#		2. Redistributions in binary form must reproduce the above copyright
#		notice, this list of conditions and the following disclaimer in the
#		documentation and/or other materials provided with the distribution.
#		
#		3. Neither the name of the copyright holder nor the names of its
#		contributors may be used to endorse or promote products derived from
#		this software without specific prior written permission.
#		
#		THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
#		"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
#		LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
#		A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
#		HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
#		SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
#		LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
#		DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
#		THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
#		(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
#		OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#============================================================================
#		Imports
#----------------------------------------------------------------------------
require 'getoptlong';
require 'rubygems';
require 'hpricot';





#============================================================================
#		cleanedText : Clean the XML text.
#----------------------------------------------------------------------------
def cleanedText(theText, cleanAll)

	# Get the state we need
	typoWords   = {	"acul-de-sac"	=> "a cul-de-sac",
					"cul-de sac"	=> "cul-de-sac",
					"wardsthen"		=> "wards then",
					"wardsto"		=> "wards to",
					"southa"		=> "south a" };
	
	rammedWords = [	"PRIVATE", "AVENUE", "COTTAGES", "PLACE", "CLOSE", "WALK",
					"DRIVE", "BRAE", "COURT", "RIGG", "LANE", "GROVE", "GLEBE",
					"GATE", "GREEN", "PARKWAY", "STREET", "extending" ];

	recaseWords = [ "from", "at", "to" ];



	# Fix Hpricot output
	theText.gsub!("&#10", "");



	# Fix Edinburgh Council formatting
	if (cleanAll)
		# Fix whitespace
		theText.strip!;
		theText.gsub!(/\s\s+/,					" ");
		theText.gsub!(/([a-z0-9])([A-Z])/,		"\\1 \\2");
		theText.gsub!(/([.,])(\w)/,				"\\1 \\2");

		theText.gsub!(/([A-Z])From/,			"\\1 from");
		theText.gsub!(/([A-Z])Part/,			"\\1 part");
		theText.gsub!(/([A-Z])Not/,				"\\1 not");

		rammedWords.each do |theWord|
			theText.gsub!(/(\w)(#{theWord}\b)/,	"\\1 \\2");
			theText.gsub!(/\b(#{theWord})(\w)/,	"\\1 \\2");
		end


		# Fix case
		recaseWords.each do |theWord|
			theText.gsub!(/ #{theWord} /i, " #{theWord} ");
		end



		# Fix typos
		typoWords.each_pair do |badWord, goodWord|
			theText.gsub!(/#{badWord}/, "#{goodWord}");
		end
	end

	return(theText);

end





#============================================================================
#		xml2Csv : Convert the XML spreadsheets to CSV.
#----------------------------------------------------------------------------
def xml2Csv()

	# Parse the XML
	theTable = Array.new();

	ARGV.each do |thePath|

		theFile = File.new(thePath);
		theDoc  = open(thePath) { |f| Hpricot(f) }

		theDoc.search("/Workbook/Worksheet/Table/Row").each do |theRow|
		
			theEntry = Array.new(4, "");

			theRow.search("/Cell").each do |theCell|
				theColumn = theCell["ss:index"].to_i - 1;
				theText   = theCell.innerText;

				theText             = cleanedText(theText, theColumn == 3);
				theEntry[theColumn] = theText;
			end
		
			theTable << theEntry if (theEntry[0] != "Name");

		end
	end



	# Generate the CSV
	puts "\"Name\",\"Locality\",\"Street Adoption Status\",\"Property Notice Description\"";

	theTable.each do |theRow|
		puts "\"#{theRow[0]}\",\"#{theRow[1]}\",\"#{theRow[2]}\",\"#{theRow[3]}\"";
	end

end





#============================================================================
#		showHelp : Show the help.
#----------------------------------------------------------------------------
def showHelp()

	puts "edroads: process Edinburgh Council road names"
	puts "";
	puts "  edroads.rb --xml2csv file1.xls [file2.xls] [fileN.xls]"
	puts "";
	puts "    Convert xls files to csv output. xls files should be converted from"
	puts "    Edinburgh Council .pdfs using the pdftoexcelonline.com converter.";
	puts "";

end





#============================================================================
#		edroads : Edinburgh Council road processor
#----------------------------------------------------------------------------
def edroads()


	# Process the parameters
	theParams = GetoptLong.new(	[ '--help',    '-h', GetoptLong::NO_ARGUMENT ],
								[ '--xml2csv', '-c', GetoptLong::NO_ARGUMENT ] );

	theParams.each do |theFlag, theArg|
		case theFlag
			when '--xml2csv'
				xml2Csv();
				return;
		end
	end



	# Fall through to help
	showHelp();

end





#============================================================================
#		Entry point
#----------------------------------------------------------------------------
edroads();
