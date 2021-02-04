Deleting Empty Outlook Appointments
===================================

Until recently, I've always kept my main set of contacts on my personal mobile device and used Exchange ActiveSync to synchronize it with my mobile device. When I joined Forrester, I decided to keep my work stuff separate from my personal stuff. So, rather than have all of my personal contacts available to me while at work, I let my work account sync my work PIM data to my device through its normal means and used a separate process to sync my person PIM data to my device as well.

Since I carry an Android device, I looked around and found an open source solution called Go Contact Sync Mod (http://googlesyncmod.sourceforge.net/) that works pretty well. I could have the sync run automatically, but for now I'm running it manually. One of the things I noticed when I run it is that it was showing an error indicating that it had found empty calendar entries in Outlook. I didn't know I had empty calendar entries nor did I know you could create empty calendar entries, but apparently my pst file had a bunch of them.
 
I looked around for a while and didn't find a solution for deleting those empty entries, and I couldn't 'see' them in Outlook, so I decided to write some code to solve the problem. I'd been helping a friend with some Outlook automation, so I had the basis of what I needed to make this work. I'm a big Delphi developer, although I'm pretty upset with Embarcadero right now, so it gave me a chance to dust off my Delphi skills and write some code.

The code that accesses the Outlook calendar from Delphi came from [http://www.scalabium.com/faq/dct0128.htm](http://www.scalabium.com/faq/dct0128.htm). 

The code that accesses Outlook items came from [http://www.scalabium.com/faq/dct0121.htm](http://www.scalabium.com/faq/dct0121.htm). 

***

You can find information on many different topics on my [personal blog](http://www.johnwargo.com). Learn about all of my publications at [John Wargo Books](http://www.johnwargobooks.com).

If you find this code useful and feel like thanking me for providing it, please consider <a href="https://www.buymeacoffee.com/johnwargo" target="_blank">Buying Me a Coffee</a>, or making a purchase from [my Amazon Wish List](https://amzn.com/w/1WI6AAUKPT5P9).
