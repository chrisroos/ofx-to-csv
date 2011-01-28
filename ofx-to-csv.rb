#!/usr/bin/env ruby

require 'rubygems'
require 'bundler/setup'

require 'ofx-parser'
require 'fastercsv'
require 'active_support/inflector'

ofx_file = ARGV.shift
unless ofx_file and File.exists?(ofx_file)
  puts "Usage: #{File.basename(__FILE__)} /path/to/statement.ofx"
  exit 1
end

ofx = OfxParser::OfxParser.parse(open(ofx_file))

ofx.bank_account.statement.transactions.each do |transaction|
  date  = transaction.date.strftime('%Y-%m-%d')
  payee = ActiveSupport::Inflector.titleize(transaction.payee)
  memo  = ActiveSupport::Inflector.titleize(transaction.memo)
  type  = ActiveSupport::Inflector.titleize(transaction.type)
  
  puts [date, transaction.fit_id, payee, memo, type, transaction.amount].to_csv
end