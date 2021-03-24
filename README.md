# abodian_scripts
Various scripts for editing & creating files for Mozaik.
(editing materials, editing material templates, adding new materials and creating new material templates with those materials, printing special reports, renaming backups, etc.)

Mozaik is a 3D modeling software that generates drawings for our clients and part lists and CNC machine code for our production team to machine all of our parts and build our products.

Mozaik's software has some limitations that we have needed to circumvent through writing our own scripts.

One of the scripts we needed was the ability to copy every door material in our system in order to create a duplicate material with the same name that had the text "BSAW" at the
end of the name and some modifications to the trim dimensions for the edges of the material. This function allowed us to assign this new material to every material template
(roughly 700 of them) for specific parts we want to send to the Beam Saw instead of our Nested Router CNC. I was able to assign these new materials to each template using
another script. Both of these functions can be found in the material_template_update.py file.

Another challenge we've been working to overcome is that the software does not have a report that lists all products that had special notes for the shop.
Our shop employees have been failing to read the notes associated with some products before shipping, so we've had a decent amount of service work to replace products.
Previously, these special notes were only found on the product labels, which our team only adds to the product once it's complete and ready to be wrapped.
We needed a report they could look at prior to and during assembly that would help ensure these notes were followed prior to completing assembly.
The purpose of the product_list_note.py file is to read each of the .des files inside a project directory and pull out all of the products that have special notes.
Then I am creating an excel spreadsheet that lists each of these products inside their respective rooms with the special note next to each product.
This should save our company thousands of dollars in re-work.
