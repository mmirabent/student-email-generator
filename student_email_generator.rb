require "rubygems"
require "spreadsheet"
require "CSV"

# Display a message and wait for input. This is important so the error message
# can be read before the window closes
def exit_with_msg(msg)
  $stderr.puts msg + "\nPress enter to continue...\n"
  $stdin.getc
  exit false
end

if ARGV.length < 1
  exit_with_msg("Must specify an input file")
end

# Get the current school year from the user, abort if its in the wrong format
puts "Please enter the current school year (Ex: 2013-2014)"
input = $stdin.gets.chomp
if input =~ /^\d{4}-\d{4}$/
  seniorGradYear = input.slice(-4..-1).to_i
else
  exit_with_msg("Incorrect school year format")
end

# get the absolute path for the first argument which should be a file name
dataFile = File.expand_path(ARGV.first)
outFile   = File.dirname(dataFile) + "/" + "Bulk Upload.csv"
emailDomain = "@pacespartans.com"


# Check if input file exists
if !File.exists?(dataFile)
  exit_with_msg "File %s not found" % dataFile
end


# Check that the output file is writable
if !File.writable?(dataFile)
  exit_with_msg "could not write to output file %s" % outFile
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

# The headerHash should contain all of these keys
headerArray = [:first_name, :last_name, :birth_date, :apid]
if !headerArray.all? { |key| headerHash.member?(key) }
  exit_with_msg "The input file %s is malformed, it should have a header with the values %s" % [dataFile, headerArray.to_s]
end

# setup the failedBirthdate array
failedBirthdate = []

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

    if birth_date.respond_to?(:strftime)
      password    = birth_date.strftime("%m-%d-%Y")
    else
      password    = ""
      failedBirthdate.push apid
    end

    grad_year   = (seniorGradYear % 1000 + 12 - apid / 1000).to_s
    student_no  = "%03d" % (apid % 1000)

    email = firstName[0] + sanitize(lastName) + grad_year + student_no + emailDomain

    outCSV << [email, firstName, lastName, password]
  end
end

# Print error if any students had bad birth dates
if !failedBirthdate.empty?
  puts "Students with the following APIDs had invalid birth dates"
  failedBirthdate.each do |apid|
    puts apid
  end
end
