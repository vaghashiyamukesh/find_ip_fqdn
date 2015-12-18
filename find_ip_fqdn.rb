require 'net/ping'
require 'ipaddr'
require 'writeexcel'
require 'socket'
require "simple-spreadsheet"

def is_fqdn?(fqdn)
  if fqdn == 'Not Listed'
    false
  else
    true
  end
end

def is_ip?(ip)
  !!IPAddr.new(ip) rescue false
end

def up?(fqdn,ip)
  check = Net::Ping::External.new(fqdn).ping? rescue false
  unless check
    check = Net::Ping::External.new(ip).ping? rescue false
  end
check
end

def getfp(host)
  server_record = {}
  domain = ["abc.yahoo.com","abc.google.net"]
  (0..domain.length-1).each do |j|
    fq = "#{host}.#{domain[j]}"
    ip = IPSocket.getaddress(fq) rescue false
    if is_ip?(ip)
      server_record[:fqdn] = fq
      server_record[:ip] = ip
    end
  end
  server_record
end

def is_validate?(fqdn,ip,host)
  server_record = {}
  server_record[:fqdn] = fqdn
  server_record[:ip] = ip
  if !is_fqdn?(fqdn)
    server_record = getfp(host)
  end
  server_record
end

workbook   = WriteExcel.new("output.xls")
worksheet  = workbook.add_worksheet('output')
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 20)
worksheet.set_column('E:E', 20)

worksheet.write('A1', "Itrc Fqdn")
worksheet.write('B1', "itrc_host")
worksheet.write('C1', "Itrc Serial Nbr")
worksheet.write('D1', "itrc_ipaddress")
worksheet.write('E1', "Status")
count = 2

file =  ARGV.first
s = SimpleSpreadsheet::Workbook.read(file)
s.selected_sheet = s.sheets.first
s.first_row.upto(s.last_row) do |line|
  _fqdn = s.cell(line, 1,1)
  _host = s.cell(line, 2,1)
  _serial = s.cell(line,3,1)
  _ip = s.cell(line,4,1)
  server_record = is_validate?(_fqdn,_ip,_host)
  fqdn = server_record[:fqdn]
  ip = server_record[:ip]
  if is_ip?(ip)
    status = up?(fqdn,ip)
    if status
      worksheet.write("A#{count}", "#{fqdn}")
      worksheet.write("B#{count}", "#{_host}")
      worksheet.write("C#{count}", "#{_serial}")
      worksheet.write("D#{count}", "#{ip}")
      worksheet.write("E#{count}", "Active")

    else
      worksheet.write("A#{count}", "#{fqdn}")
      worksheet.write("B#{count}", "#{_host}")
      worksheet.write("C#{count}", "#{_serial}")
      worksheet.write("D#{count}", "#{ip}")
      worksheet.write("E#{count}", "Dead")
    end
    count += 1
  end
end
workbook.close