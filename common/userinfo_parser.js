/*
 * (c) Copyright Ascensio System SIA 2010-2021
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
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
