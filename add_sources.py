from pocky import pocky
from pathlib import Path
import feedparser
import datetime
from datetime import timedelta

key_path = Path('C:/Users/Andrew/OneDrive/Documents/Code Keys/Pocket/keys.json')
p = pocky(key_path)

#Get previously loaded stories....
history = p.load_history()

for f in p.get_sources().keys():
    source = p.get_sources()[f]
    feed_url = source['url']
    tags = source['tags']
    feed = feedparser.parse(feed_url)
    #print(feed)

    #Don't import articles more than 30 days old.
    oldest_article = datetime.datetime.now() - timedelta(days=30)

    #Loop through every tiem...
    for i in feed['items']:
        #Don't add if it's already been added.
        if(i['link'] in history):
            print('Already added. Skipping: ', i['title'], i['link'])
            continue

        #Don't import articles that are too old.
        published_date = datetime.datetime(i['published_parsed'][0], i['published_parsed'][1], i['published_parsed'][2])
        if(published_date < oldest_article):
            print('Skipping: Article is too old to add.')
            continue

        print('ADDING.....')
        print(i['title'])
        print(i['link'])
        p.add(i['link'], tags)
        p.add_history(i['link'])
quit()
