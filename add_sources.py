from pocky import pocky
from pathlib import Path
import feedparser

key_path = Path('C:/Users/Andrew/OneDrive/Documents/Code Keys/Pocket/keys.json')
p = pocky(key_path)

#Get previously loaded stories....
history = p.load_history()

for f in p.get_sources().keys():
    source = p.get_sources()[f]
    feed_url = source['url']
    tags = source['tags']
    feed = feedparser.parse(feed_url)
    print(feed)

    for i in feed['items']:
        #Don't add if it's already been added.
        if(i['link'] in history):
            print('Already added. Skipping: ', i['title'], i['link'])
            continue

        print('ADDING.....')
        print(i['title'])
        print(i['link'])
        p.add(i['link'], tags)
        p.add_history(i['link'])
quit()
