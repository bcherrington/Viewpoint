=begin
  This file is part of Viewpoint; the Ruby library for Microsoft Exchange Web Services.

  Copyright © 2011 Dan Wanek <dan.wanek@gmail.com>

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

  # A Generic Class for SOAP returns.
  class EwsResponse
    include Viewpoint::StringUtils

    def initialize(sax_hash)
      @resp = sax_hash
    end

    def envelope
      @resp[:envelope]
    end

    def header
      envelope[:header]
    end

    def body
      envelope[:body]
    end

    def response
      body
    end

    def response_messages
      return @response_messages if @response_messages

      @response_messages = []
      response_type = response.keys.first
      response[response_type][:response_messages].each do |full_response_message, rm|
        if full_response_message.is_a?(Hash)
          response_message_type, rm = full_response_message.first
        else
          response_message_type = full_response_message
        end
        rm_klass = class_by_name(response_message_type)
        @response_messages << rm_klass.new({response_message_type => rm})
      end
      @response_messages
    end


    private


    def class_by_name(cname)
      begin
        if(cname.instance_of? Symbol)
          cname = camel_case(cname)
        end
        Viewpoint::EWS::SOAP.const_get(cname)
      rescue NameError => e
        ResponseMessage
      end
    end

  end # EwsSoapResponse

end # Viewpoint::EWS::SOAP
