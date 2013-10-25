#!/usr/bin/env ruby

require 'bundler/setup'

require 'yaml'
require 'viewpoint'
require 'tzinfo'
require 'prowl'
require 'ruby-growl'

include Viewpoint::EWS

# disable this to show debug information
Logging.logger.root.level = :error

if File.exists? 'config'
    config = YAML::load(File.open('config'))
else
    config = {
        'EWS' => {
            'endpoint' => 'https://wmail.tabcorp.com.au/EWS/Exchange.asmx',
            'username' => 'username',
            'password' => 'password',
            'ssl_verification' => true
        },
        'proxy' => {
            'host' => 'localhost',
            'port' => 3128
        },
        'prowl' => {
            'enabled' => true, 
            'apikey' => 'prowl-api-key',
            'application' => 'Outlook'
        },
        'growl' => {
            'enabled' => true,
            'host' => 'localhost',
            'source' => 'ruby-growl',
            'application' => 'Outlook'
        },
        'monitor_list' => [
            'Inbox',
            'new relic',
            'reviewboard'
        ],
        'check_interval' => 5,
        'timezone' => 'Australia/Melbourne',
        'mark_read' => true # mark email as read on notify?
    }

    puts "Seems first time running -- please update generated config file and re-run"
    File.open('config', 'w') { |f| f.write config.to_yaml }
    exit
end

if config['EWS']['ssl_verification']
    cli = Viewpoint::EWSClient.new config['EWS']['endpoint'], config['EWS']['username'], config['EWS']['password']
else
    puts "Warning: SSL certificate verification is disabled"
    cli = Viewpoint::EWSClient.new config['EWS']['endpoint'], config['EWS']['username'], config['EWS']['password'], http_opts: {ssl_verify_mode: 0}
end

tz = TZInfo::Timezone.get(config['timezone'])

# array of message IDs that we've already notified about during this session
notified = []

if config['prowl']['enabled']
    prowl = Prowl.new :apikey => config['prowl']['apikey'], :application => config['prowl']['application']
    prowl.set_proxy config['proxy']['host'], config['proxy']['port']
end

if config['growl']['enabled']
    growl = Growl.new config['growl']['host'], config['growl']['source']
    growl.add_notification config['growl']['application']
end

puts "Retrieving folder IDs..."
folder_id_list = config['monitor_list'].collect { |display_name| cli.find_by_name(display_name).first.id }

while true
    begin
        puts "Checking on #{Time.now}..."

        folder_id_list.each do |folder_id|
            current_folder = cli.get_folder(folder_id)

            if current_folder.unread_count > 0
                current_folder.unread_messages.each do |message|
                    unless notified.include?(message.id)
                        datetime = tz.utc_to_local(message.date_time_sent).strftime('%d/%m/%Y %H:%I')
                        desc = "From: #{message.from.name} on #{datetime}\n#{message.subject}"
                        puts desc
                        message.mark_read! if config['mark_read']
                        
                        # growl desktop notification
                        if growl
                            growl.notify "Outlook", current_folder.display_name, desc
                            puts "growl (desktop) sent"
                        end

                        # prowl mobile notification
                        if prowl
                            prowl.add :event => current_folder.display_name,
                                      :description => desc
                            puts "prowl (mobile) sent"
                        end

                        notified << message.id
                    end
                end
            end
        end
    rescue Exception => ex
        puts "Something wrong happens: #{ex.message}"
    end

    sleep config['check_interval'].to_i * 60 # 5 minutes
end
