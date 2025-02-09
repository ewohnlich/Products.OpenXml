# -*- coding: utf-8 -*-
# $Id$
"""Setup handlers for OpenXml"""

from Products.CMFCore.utils import getToolByName
from Products.OpenXml import logger

from . import config


def isNotCurrentProfile(context):
    return context.readDataFile("openxml_marker.txt") is None


def setupOpenXml(context):
    """Add MS Office 2007 content types to MimetypesRegistry"""
    if isNotCurrentProfile(context): return

    site = context.getSite()

    # Adding our file types to MTR
    mtr = getToolByName(site, 'mimetypes_registry')
    for mt_dict in config.office_mimetypes:
        main_mt = mt_dict['mimetypes'][0]
        mt_name = mt_dict['name']
        existing_mt = mtr.lookup(main_mt)
        if bool(existing_mt):
            # Already installed -- eg this could mean Plone 4's 'basic'
            # default mimetypes; should we delete what's there already?
            logger.info("%s (%s) Mime type already installed, deleting", main_mt, mt_name)
            mtr.manage_delObjects(existing_mt)
        mtr.manage_addMimeType(
            mt_name,
            mt_dict['mimetypes'],
            mt_dict['extensions'],
            mt_dict['icon_path'],
            binary=True,
            globs=mt_dict['globs'])
        logger.info("%s (%s) Mime type installed", main_mt, mt_name)

    # Registering our transform in PT
    transforms_tool = getToolByName(site, 'portal_transforms')
    if config.TRANSFORM_NAME not in transforms_tool.objectIds():
        # Not already installed
        transforms_tool.manage_addTransform(config.TRANSFORM_NAME, 'Products.OpenXml.transform')
    return


def removeOpenXml(context):
    """Removing various resources from plone site"""

    # At the moment, there's no uninstall support in GenericSetup. So
    # this is run by the old style quickinstaller uninstall handler.

    site = context
    # Removing our types from MTR
    mtr = getToolByName(site, 'mimetypes_registry')
    mt_ids = [mt_dict['mimetypes'][0] for mt_dict in config.office_mimetypes]
    mtr.manage_delObjects(mt_ids)

    # Removing our transform from PT
    transforms_tool = getToolByName(site, 'portal_transforms')
    transforms_tool.unregisterTransform(config.TRANSFORM_NAME)
    return
