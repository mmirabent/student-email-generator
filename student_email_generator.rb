require 'CSV'
require 'tempfile'
require 'roo'

tempFile  = Tempfile.new('export')
outFile   = "out.csv"

# Options for CSV::new
options = {
  :headers            => true,
  :header_converters  => :symbol,
  :converters         => :date,
  :force_quotes       => true
}

# Make the path string for the file
dataFile = File.expand_path(ARGV.first);

# Load and convert the .xls file to CSV
dataExcel = Roo::Excel.new(dataFile)
dataExcel.to_csv(tempFile)

# Instantiate the csv object for dealing with the csv file and add a header
outCSV = CSV.open(outFile, "wb", options)
outCSV << ["emailAddress", "first name", "last name", "password"]

CSV.open(outFile, "wb") do |outCSV|
  CSV.foreach(tempFile,options) do |row|
    firstName = row.field(:first_name)
    lastName = row.field(:last_name)
    outCSV << [firstName, lastName]
  end
end

tempFile.close!
