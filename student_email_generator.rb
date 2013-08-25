require 'CSV'
require 'tempfile'
require 'roo'

# If the current school year is 2013-2014, then the seniors graduate in 2014
seniorGradYear = 2014
emailDomain = "@pacespartans.com"

# People like to have funny stuff like spaces, hyphens, apostrophes, etc
# In their names. This fixes that
def sanitize(name)
  name.gsub /[\-\'\ \,\.]/, ""
end

# Temporary file for holding intermediate data. roo sucks, but I'm not about
# to re-write it for this.
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
  
  # This writes out each of the rows on our output.csv
  # Google Apps is expecting Email, First Name, Last Name, Password
  # as headers in a .csv
  # APID is defined as GGNNN where GG is the students current grade. GG is 
  # extracted by integer division. GGNNN / 1000 = GG
  # the student number is extracted by modulus GGNNN % 1000 = NNN
  # To get graduation year from grade level, take the year the current seniors
  # will graduate at, add 12 and subtract the current grade level
  CSV.foreach(tempFile,options) do |row|
    firstName   = row.field(:first_name)
    lastName    = row.field(:last_name)
    password    = row.field(:birth_date)
    apid        = row.field(:apid).to_i
    grad_year   = (seniorGradYear % 1000 + 12 - apid / 1000).to_s
    student_no  = "%03d" % (apid % 1000)
    
    email = firstName[0] + sanitize(lastName) + grad_year + student_no + emailDomain
    
    outCSV << [email, firstName, lastName, password]
  end
end

tempFile.close!
