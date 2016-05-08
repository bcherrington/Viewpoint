=begin
  This file is part of Viewpoint; the Ruby library for Microsoft Exchange Web Services.

  Copyright © 2015 Tatiana Kudiyarova <tatiana.kudiyarova@gmail.com>

  Licensed under the Apache License, Version 2.0 (the "License");
  you may not use this file except in compliance with the License.
  You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.
=end

module Viewpoint::EWS::SOAP
  class GetStreamingEventsResponseMessage < ResponseMessage
    include Viewpoint::StringUtils

    def connection_status
      safe_hash_access message, [:elems, :connection_status, :text]
    end

    def notifications
      return @notifications if @notifications

      n = safe_hash_access message, [:elems, :notifications, :elems]
      @notifications = n.nil? ? [] : parse_notifications(n)
    end

    def error_subscription_ids
      e = safe_hash_access message, [:elems, :error_subscription_ids, :elems]
      e.nil? ? [] : e.map {|m| safe_hash_access(m, [:subscription_id, :text])}
    end

    private

    def parse_notifications(notifications)
      notifications.collect do |notification|
        type = notification.keys.first
        klass = Viewpoint::EWS::Types.const_get(camel_case(type))
        klass.new(nil, notification[type])
      end
    end

  end # GetStreamingEventsResponseMessage
end # Viewpoint::EWS::SOAP
