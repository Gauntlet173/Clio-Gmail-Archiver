# Clio-Gmail-Archiver
A Google Apps Script using a spreadsheet to file email as communications in Clio.

This software is unstable.  Don't use it unless you are satisfied you know better what you are doing than the person who
wrote the software.

Even in the best-case-scenario, there is a persistent issue that while the script is running, it will only archive emails
that have had a label assigned to them since the last time an email with that label was archived.

This means that if you assign labels to emails in an order other than the order in which you received the emails, the script
may find the newer emails with the label before it finds the older emails with that label, and it will therefore never
archive the older emails.

This can be worked around by only assigning labels to the oldest emails first, or by turning the script off while you might be assigning labels in another order.

About Me
My name is Jason Morris, I'm a lawyer working in Sherwood Park, Alberta.
