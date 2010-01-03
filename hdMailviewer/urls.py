#!/usr/bin/env python
# encoding: utf-8

"""Definition der URL-View Mappings."""

from django.conf.urls.defaults import *
from django.contrib.admin.views.decorators import staff_member_required as staff
from hudjango.auth.decorators import require_login
from intern.views import search, newsearch

urlpatterns = patterns('',
    (r'^attachmentarchive/(?P<messagekey>.+)/attachment/(?P<attachmentkey>.+)', 'intern.views.attachmentarchive_attachment'),
    (r'^attachmentarchive/(?P<messagekey>.+)/', 'intern.views.attachmentarchive_message'),
    (r'^attachmentarchive/', 'intern.views.attachmentarchive_index')
)
