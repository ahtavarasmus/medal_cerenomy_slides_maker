this is the code of a gui program that creates powerpoint presentation from a csv file containing results of the competition.
it creates presentation that starts with the category and next slide starts playing the winners national anthem and top three flags start raising from the bottom, then next category same thing.

<img width="867" alt="kuva" src="https://github.com/ahtavarasmus/medal_cerenomy_slides_maker/assets/76125307/6eb3a901-10f7-4566-8138-39916b115f69">


https://github.com/ahtavarasmus/medal_cerenomy_slides_maker/assets/76125307/461bffa7-7932-4b66-a5c8-d22b935e81cb

if you are interested about my 12h struggle to create this haha:
- created code with python-pptx and imageio that reads csv and makes rising flags with 100 different slides each has flags slightly higher since you couldn't do animations with pptx library
- realized i could not modify the transitions with pptx which were needed for those 100 slides to make them move quickly to create animation
- tried to search libraries that can do that, couldn't, but found that applescript could do that. wrote applescript for it, but while trying to get it working learned it would not work for other people to use the script since they are not on mac lol.
- then found aspose.slides library that can modify transitions and add audio and modify playback settings. tried to get that working with mac - couldn't, it needs .net but didn't somehow get it installed correctly(tried 2h) then found asposeslidescloud struggled with that one for an hour, wrote the scripts for it, found out cloud version can't add audio.. lol. - then opened my windows pc since it has .net installed default, tried again with aspose.slides and found out you need a license for aspose.slides:D. it costs 1000e HAHA, luckily there was 30day free trial for freelancer business(it required custom non free email, which i was lucky to have bought month ago lol).
- I got aspose.slides to work and also the whole python script on windows
- created a pyqt6 app around it so it can be used with mouse
- tried to package it with pyinstaller, but after almost 2 hours couldn't get the exe working - quit
- tried to install aspose.slides on my schools remote linux, i was able to so thought about creating web app
- went to buy server from digital ocean - i got my account somehow locked before even getting to buy the server..
- went to buy server from contabo - was able to buy ubuntu
- configured it and got flask app running on the ip address, but it gave error No usable version of libssl was foundAborted
- tried to fix that for 2 hours with installing python openssl from source and had much trouble linking python to the newer openssl version
- after succeeding with linking same libssl error continued - gave up.
- tried to package my original pyqt6 app with py2exe - couldn't. my app needs files to work so i had trouble getting them included.
- tried again with pyinstaller and shoutout to chatgpt modified the spec file so that my program exe finally worked.
- made the ui better and added couple settings for generating the slides
- done:D phew
