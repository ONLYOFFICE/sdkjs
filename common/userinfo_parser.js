/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

(function(window) {
    var parse = false,
        separator = String.fromCharCode(160),
        username = '',
        _reviewPermissions,
        reviewGroups,
        commentGroups,
        userInfoGroups;

    var _intersection = function (arr1, arr2) {
        if (arr1 && arr2) {
            for (var i=0; i<arr2.length; i++) {
                if (arr1.indexOf(arr2[i])>-1)
                    return true;
            }
        }
        return false;
    };

    var UserInfoParser  = {
        setParser: function(value) {
            parse = !!value;
        },

        setCurrentName: function(name) {
            username = name;
            _reviewPermissions && this.setReviewPermissions(null, _reviewPermissions); // old version of review permissions
        },

        getCurrentName: function() {
            return username;
        },

        getSeparator: function() {
            return separator;
        },

        getParsedName: function(username) {
            return (parse && username) ? username.substring(username.indexOf(separator)+1) : username;
        },

        getParsedGroups: function(username) {
            if (parse && username) {
                var idx = username.indexOf(separator),
                    groups = (idx>-1) ? username.substring(0, idx).split(',') : [];
                for (var i=0; i<groups.length; i++)
                    groups[i] = groups[i].trim();
                return groups;
            }
        },

        setReviewPermissions: function(groups, permissions) {
            if (groups) {
                if  (typeof groups == 'object' && groups.length>=0)
                    reviewGroups = groups;
            } else if (permissions) { // old version of review permissions
                var arr = [],
                    arrgroups  =  this.getParsedGroups(username);
                arrgroups && arrgroups.forEach(function(group) {
                    var item = permissions[group.trim()];
                    item && (arr = arr.concat(item));
                });
                reviewGroups = arr;
                _reviewPermissions = permissions;
            }
        },

        setCommentPermissions: function(groups) {
            if (groups && typeof groups == 'object') {
                commentGroups = {
                    view: (typeof groups["view"] == 'object' && groups["view"].length>=0) ? groups["view"] : null,
                    edit: (typeof groups["edit"] == 'object' && groups["edit"].length>=0) ? groups["edit"] : null,
                    remove: (typeof groups["remove"] == 'object' && groups["remove"].length>=0) ? groups["remove"] : null
                };
            }
        },

        setUserInfoPermissions: function(groups) {
            if (groups) {
                if  (typeof groups == 'object' && groups.length>=0)
                    userInfoGroups = groups;
            }
        },

        getCommentPermissions: function(permission) {
            if (parse && commentGroups) {
                return commentGroups[permission];
            }
        },

        canEditReview: function(username) {
            if (!parse || !reviewGroups) return true;

            var groups = this.getParsedGroups(username);
            groups && (groups.length==0) && (groups = [""]);
            return _intersection(reviewGroups, groups);
        },

        canViewComment: function(username) {
            if (!parse || !commentGroups || !commentGroups.view) return true;

            var groups = this.getParsedGroups(username);
            groups && (groups.length==0) && (groups = [""]);
            return _intersection(commentGroups.view, groups);
        },

        canEditComment: function(username) {
            if (!parse || !commentGroups || !commentGroups.edit) return true;

            var groups = this.getParsedGroups(username);
            groups && (groups.length==0) && (groups = [""]);
            return _intersection(commentGroups.edit, groups);
        },

        canDeleteComment: function(username) {
            if (!parse || !commentGroups || !commentGroups.remove) return true;

            var groups = this.getParsedGroups(username);
            groups && (groups.length==0) && (groups = [""]);
            return _intersection(commentGroups.remove, groups);
        },

        isUserVisible: function(username) {
            if (!parse || !userInfoGroups) return true;

            var groups = this.getParsedGroups(username);
            groups && (groups.length==0) && (groups = [""]);
            return _intersection(userInfoGroups, groups);
        }
    }

    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon']['UserInfoParser'] = window['AscCommon'].UserInfoParser = UserInfoParser;
    UserInfoParser['setParser'] = UserInfoParser.setParser;
    UserInfoParser['setCurrentName'] = UserInfoParser.setCurrentName;
    UserInfoParser['getCurrentName'] = UserInfoParser.getCurrentName;
    UserInfoParser['getSeparator'] = UserInfoParser.getSeparator;
    UserInfoParser['getParsedName'] = UserInfoParser.getParsedName;
    UserInfoParser['setReviewPermissions'] = UserInfoParser.setReviewPermissions;
    UserInfoParser['setCommentPermissions'] = UserInfoParser.setCommentPermissions;
    UserInfoParser['setUserInfoPermissions'] = UserInfoParser.setUserInfoPermissions;
    UserInfoParser['canEditReview'] = UserInfoParser.canEditReview;
    UserInfoParser['canViewComment'] = UserInfoParser.canViewComment;
    UserInfoParser['canEditComment'] = UserInfoParser.canEditComment;
    UserInfoParser['canDeleteComment'] = UserInfoParser.canDeleteComment;
    UserInfoParser['isUserVisible'] = UserInfoParser.isUserVisible;
    UserInfoParser['getParsedGroups'] = UserInfoParser.getParsedGroups;
    UserInfoParser['getCommentPermissions'] = UserInfoParser.getCommentPermissions;

})(window);
