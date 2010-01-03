#!/usr/bin/env python
# encoding: utf-8

"""Views für vermischte interne Seiten in der HUDORA Webapplikation."""

import operator

from django.template import RequestContext
from django.shortcuts import render_to_response
from django.db import models
from django.http import HttpResponseRedirect, HttpResponse
from django.contrib.admin.views.decorators import staff_member_required
import httplib2
import urllib
import feedparser
from produktpass.models import Product
import couchdb
import urlparse


def attachmentarchive_index(request):
    start = int(request.GET.get('start', '0'))
    perpage = 1000
    server = couchdb.client.Server('http://couchdb1.local.hudora.biz:5984/')
    db = server['attachments']

    query = request.GET.get('query', '')
    if query and (query.strip() in db):
        return HttpResponseRedirect('./%s/' % urllib.quote(query.strip()))

    results = db.view('_all_docs', limit=perpage, skip=start)
    return render_to_response('hdMailviewer/attachmentarchive_index.html',
                              {'title': 'Urbersicht archivierte Attachments', 
                               # wir erzwingen volle tausenderschritte
                               'nextstart': int((start + perpage) / 1000) * 1000,
                               'results': results, 'query': query.strip()},
                              context_instance=RequestContext(request))

def attachmentarchive_message(request, messagekey):
    server = couchdb.client.Server('http://couchdb1.local.hudora.biz:5984/')
    db = server['attachments']
    doc = db[messagekey]
    doc['attachments'] = doc['_attachments'] # needed for django Template engine
    return render_to_response('hdMailviewer/attachmentarchive_message.html',
                              {'title': 'archivierte Attachments: %s (%s)' % (doc.get('subject'), doc.get('date')),
                               'key': messagekey,
                               'doc': doc},
                              context_instance=RequestContext(request))

def attachmentarchive_attachment(request, messagekey, attachmentkey):
    server = couchdb.client.Server('http://couchdb1.local.hudora.biz:5984/')
    db = server['attachments']
    doc = db[messagekey]
    
    attachment = doc['_attachments'][attachmentkey]
    response = HttpResponse(mimetype=attachment['content_type'])
    response.write(db.get_attachment(doc, attachmentkey))
    return response
  