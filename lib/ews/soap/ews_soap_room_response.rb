=begin
  This file is a contribution to Viewpoint; the Ruby library for Microsoft Exchange Web Services.

  Copyright © 2013 Camille Baldock <viewpoint@camillebaldock.co.uk>

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

  # A class for roomlists SOAP returns.
  # @attr_reader [String] :message The text from the EWS element <m:ResponseCode>
  class EwsSoapRoomResponse < EwsSoapResponse

    def response_messages
      key = response.keys.first
      subresponse = response[key]
      response_class = subresponse.keys.first
      subresponse[response_class]
    end

    def roomsArray
      response[:get_rooms_response][:rooms]
    end

    def success?
      response.first[:response_class] == "Success"
    end

    private


  end # EwsSoapRoomResponse

end # Viewpoint::EWS::SOAP
