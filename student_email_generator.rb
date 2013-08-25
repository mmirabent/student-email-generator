require 'CSV'
require 'tempfile'
require 'roo'

# If the current school year is 2013-2014, then the seniors graduate in 2014
seniorGradYear = 2014
emailDomain = "@pacespartans.com"

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
CSV.open(outFile, "wb",options) do |outCSV|
  outCSV << ["email address", "first name", "last name", "password"]
  CSV.foreach(tempFile,options) do |row|
    firstName   = row.field(:first_name)
    lastName    = row.field(:last_name)
    password    = row.field(:birth_date)
    apid        = row.field(:apid).to_i
    grad_year   = (seniorGradYear % 1000 + 12 - apid / 1000).to_s
    student_no  = "%03d" % (apid % 1000)
    
    email = firstName[0] + lastName + grad_year + student_no + emailDomain
    
    outCSV << [email, firstName, lastName, password]
  end
end

tempFile.close!
