# Editor
uses quill for the text editing
custom component for the spotlight like search bar (for now a bottom preview is also added) for documents refer to modal-overlay

## work invloved 
1. this cant be used when we are using the remote language (less than 1% of the case) ~ 3 clients
2. the value is stored in the form of a  html or quill delta -- that is converted as of now
3. Need to regenrate the links and replace them in the source when the document is previewed or generated (this will increase report generation time as of now we use S3 to generate the urls)
3. need to update the docx templates to entertain the new wml format 
4. in order for a 2 way sync we need to update the original report builder also as well as PAI? because quill delta will have styles that cannot be shown in the report language at ALL. --- IMP discuss
5. This editor should also have a PAI button that goes well with the bottom preview of docs for both sides


