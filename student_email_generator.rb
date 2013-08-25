require "rubygems"
require "spreadsheet"
require "CSV"

if ARGV.length < 1
  abort("Must specify an input file")
end

# get the absolute path for the first argument which should be a file name
dataFile = File.expand_path(ARGV.first)
outFile   = File.dirname(dataFile) + "/" + "Bulk Upload.csv"
seniorGradYear = 2014
emailDomain = "@pacespartans.com"


# Check if input file exists
if !File.exists?(dataFile)
  abort "File %s not found" % dataFile
end

# Check that the output file is writable
if !File.writable?(dataFile)
  abort "could not write to output file %s" % outFile
end

# Options for CSV::new
options = {
  :headers            => true,
  :header_converters  => :symbol,
  :converters         => :date,
  :force_quotes       => true
}

# turn arbitrary strings into symbols. Downcase, replace spaces with underscores
# and finally, drop any characters that arent letters of underscores
def symbolize(str)
  str.downcase.gsub(/\ /, "_").gsub(/[^a-z_]/, "").to_sym
end

def sanitize(name)
  name.gsub /[^a-zA-Z]/, ""
end

workbook = Spreadsheet.open(dataFile)
worksheet = workbook.worksheet 0

# Find header row and prepare header hash
headerRow = worksheet.row worksheet.dimensions[0]
headerHash = Hash.new

# Process header row
headerRow.each_index do |cell_index|
  headerHash[symbolize headerRow[cell_index]] = cell_index
end

# Open the CSV file for writing
CSV.open(outFile, "wb",options) do |outCSV|
  # write out the header file
  outCSV << ["email address", "first name", "last name", "password"]

  # Begin iterating after the header row
  worksheet.each(worksheet.dimensions[0]+1) do |row|
    # This writes out each of the rows on our output.csv
    # Google Apps is expecting Email, First Name, Last Name, Password
    # as headers in a .csv
    # APID is defined as GGNNN where GG is the students current grade. GG is 
    # extracted by integer division. GGNNN / 1000 = GG
    # the student number is extracted by modulus GGNNN % 1000 = NNN
    # To get graduation year from grade level, take the year the current seniors
    # will graduate at, add 12 and subtract the current grade level
    firstName   = row[headerHash[:first_name]]
    lastName    = row[headerHash[:last_name]]
    birth_date  = row[headerHash[:birth_date]]
    apid        = row[headerHash[:apid]].to_i
    password    = birth_date.strftime("%m-%d-%Y")
    grad_year   = (seniorGradYear % 1000 + 12 - apid / 1000).to_s
    student_no  = "%03d" % (apid % 1000)

    email = firstName[0] + sanitize(lastName) + grad_year + student_no + emailDomain

    outCSV << [email, firstName, lastName, password]
  end
end
