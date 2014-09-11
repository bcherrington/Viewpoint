module Viewpoint::EWS::Types
  class ContactsFolder
    include Viewpoint::EWS
    include Viewpoint::EWS::Types
    include Viewpoint::EWS::Types::GenericFolder

    # Creates a contact
    # @param attributes [Hash] Parameters of the contact.
    # @return [Contact]
    # @see Template::Contact
    def create_item(attributes)
      template = Viewpoint::EWS::Template::Contact.new attributes
      template.saved_item_folder_id = {id: self.id, change_key: self.change_key}
      rm = ews.create_item(template.to_ews_create).response_messages.first
      if rm && rm.success?
        Contact.new ews, rm.items[:contact]
      else
        raise EwsCreateItemError, "Could not create contact in folder. #{rm.code}: #{rm.message_text}" unless rm
      end
    end
  end
end
