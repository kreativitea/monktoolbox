''' This file manages all the communication between the program and d3up.  
It also manages the configuration (since the config is the profile link.) '''

import time
import requests
from urlparse import urlparse
from requests import ConnectionError

def error_out(message, t):
    print message
    print "Closing in {} seconds...".format(t)
    time.sleep(t)
    return message


def load_profile():
    f = open('profile.cfg')
    with f:
        print 'loading profile link from profile.cfg configuration file...'
        profile_link = f.next()

    parse = urlparse(profile_link)
    if parse.netloc == 'd3up.com':
        return profile_link
    else:
        m = error_out("The profile link provided in profile.cfg is invalid! (not a d3up link)", 10)
        raise ValueError(m)


def parse_link(profile):
    ''' Goes through the provided profile link until it finds a number.
    It then adds the json tag to the list, and returns the complete json link.
    
    >>> parse_link('http://d3up.com/b/123456/somechar')
    'http://d3up.com/b/123456/json'
    '''
    splitprof = profile.split('/')
    for i, s in enumerate(splitprof):
        try:
            int(s)
            link = '/'.join(splitprof[:i+1] + ['json'])
            print 'profile link validated! {}'.format(link)
            return link
        except ValueError:
            pass


def download_data():
    ''' Downloads the data from the profile link specified in profile.cfg.  
    Connects using requests.'''
    profile = parse_link(load_profile())
    print 'downloading data now...'
    try:
        r = requests.get(profile)
        if r.status_code == 200:
            print 'download complete!'
            # sometimes you feel like a method, sometimes you dont.
            try:
                r.json.get('testdata', None)
                return r.json
            except:
                return r.json()
            
        else:
            m = error_out("Could not connect to d3up!", 10)
            raise IOError(m)
    except ConnectionError:
       m = error_out("Could not connect to the internet!", 10)
       raise ConnectionError