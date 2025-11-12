/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
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

    // $(function() {
    //     window['Asc'] = window['Asc'] || {};

    
    // }
window['Asc'].document1Base64 = "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz4KPGN4OmNoYXJ0U3BhY2UgeG1sbnM6YT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL2RyYXdpbmdtbC8yMDA2L21haW4iIHhtbG5zOnI9Imh0dHA6Ly9zY2hlbWFzLm9wZW54bWxmb3JtYXRzLm9yZy9vZmZpY2VEb2N1bWVudC8yMDA2L3JlbGF0aW9uc2hpcHMiIHhtbG5zOmN4PSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS9kcmF3aW5nLzIwMTQvY2hhcnRleCI+CiAgICA8Y3g6Y2hhcnREYXRhPgogICAgICAgIDxjeDpleHRlcm5hbERhdGEgaWQ9InN0cmluZyIgYXV0b1VwZGF0ZT0iMSIgLz4KICAgICAgICA8Y3g6ZGF0YSBpZD0iMjkxOSI+CiAgICAgICAgICAgIDxjeDpudW1EaW0gdHlwZT0iY29sb3JWYWwiPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSI4MTY5IiBmb3JtYXRDb2RlPSJzdHJpbmciPgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjUyNjMiPjQzNDAxODAwMDAwMDAwLjI8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjY4NzYiPjQ2ODE1MDAwMDAwMDAuMjQ8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjgxMSI+LTMxMjg3NTk5OTk5OTk5Ljg8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjQ1MzMiIGZvcm1hdENvZGU9InN0cmluZyIgbmFtZT0ic3RyaW5nIj4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIzMDQiPi0xMTg5OTU5OTk5OTk5OS44PC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIzNDMxIj40NTc0NTUwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI1MDUxIj4yMTU5NzIwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICA8L2N4Om51bURpbT4KICAgICAgICAgICAgPGN4Om51bURpbSB0eXBlPSJzaXplIj4KICAgICAgICAgICAgICAgIDxjeDpmIGRpcj0iY29sIj5zdHJpbmc8L2N4OmY+CiAgICAgICAgICAgICAgICA8Y3g6bmY+c3RyaW5nPC9jeDpuZj4KICAgICAgICAgICAgPC9jeDpudW1EaW0+CiAgICAgICAgICAgIDxjeDpzdHJEaW0gdHlwZT0iZW50aXR5SWQiPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSI4OTU0IiBuYW1lPSJzdHJpbmciPgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjE3MDEiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNjQ0MSI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI4MjY4Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjIxOTIiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICA8L2N4Omx2bD4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iMTAwMDAiPgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjgwOTUiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iMTAwOSI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI3MjkxIj5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgIDwvY3g6c3RyRGltPgogICAgICAgICAgICA8Y3g6c3RyRGltIHR5cGU9ImNhdCI+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjY3MTAiIG5hbWU9InN0cmluZyI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNzkzNiI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI3NjQ1Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjkzNzUiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iMjk1MCI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSI1MjI4Ij4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI4MjA2Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgIDwvY3g6c3RyRGltPgogICAgICAgICAgICA8Y3g6ZXh0THN0PgogICAgICAgICAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgICAgICAgICAgICAgIDxjeDpleHQgLz4KICAgICAgICAgICAgICAgIDxjeDpleHQgLz4KICAgICAgICAgICAgPC9jeDpleHRMc3Q+CiAgICAgICAgPC9jeDpkYXRhPgogICAgICAgIDxjeDpkYXRhIGlkPSIxMDcwIj4KICAgICAgICAgICAgPGN4Om51bURpbSB0eXBlPSJ5Ij4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iMzYwOSIgbmFtZT0ic3RyaW5nIj4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI5NTg2Ij40ODg4MTIwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIzODcxIj4yMjc0MDAwMDAwMDAwLjIzPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIyODAzIj42NTc2MDAwMDAwMDAuMjM1PC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICA8L2N4Om51bURpbT4KICAgICAgICAgICAgPGN4OnN0ckRpbSB0eXBlPSJjb2xvclN0ciI+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9Ijc0MjMiIG5hbWU9InN0cmluZyIgLz4KICAgICAgICAgICAgPC9jeDpzdHJEaW0+CiAgICAgICAgICAgIDxjeDpzdHJEaW0gdHlwZT0iZW50aXR5SWQiPgogICAgICAgICAgICAgICAgPGN4OmYgZGlyPSJjb2wiPnN0cmluZzwvY3g6Zj4KICAgICAgICAgICAgICAgIDxjeDpuZiBkaXI9ImNvbCI+c3RyaW5nPC9jeDpuZj4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iNTY4MSI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNzc3Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjUxMzAiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iMTU1OCI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI4MTM3Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjY0MjUiIG5hbWU9InN0cmluZyI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iOTk0NSI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIzMDE2Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjY0NSI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSI0NjU3IiBuYW1lPSJzdHJpbmciIC8+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjIzMDIiIG5hbWU9InN0cmluZyI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iMTk4Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjM4ODgiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICA8L2N4Omx2bD4KICAgICAgICAgICAgPC9jeDpzdHJEaW0+CiAgICAgICAgICAgIDxjeDpzdHJEaW0gdHlwZT0iY2F0Ij4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iMTY5Ij4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI1ODc4Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9Ijc3MjIiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iMzI0MiI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI0MjA0Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9Ijk3NDkiPgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjEzMjMiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICA8L2N4Omx2bD4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iODk3MyI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNzg1MCI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI1MTQwIj5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjM1MjQiPgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjEyMTIiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iMTk0OSI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI0OTY1Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjU4MTMiIG5hbWU9InN0cmluZyI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNzg0MSI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICA8L2N4OnN0ckRpbT4KICAgICAgICAgICAgPGN4Om51bURpbSB0eXBlPSJzaXplIj4KICAgICAgICAgICAgICAgIDxjeDpmPnN0cmluZzwvY3g6Zj4KICAgICAgICAgICAgICAgIDxjeDpuZj5zdHJpbmc8L2N4Om5mPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSIyNjQ4IiBmb3JtYXRDb2RlPSJzdHJpbmciIG5hbWU9InN0cmluZyI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNjM3OCI+LTE4MDI4Nzk5OTk5OTk5Ljg8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjM2OTAiPi0yNTU0Njg5OTk5OTk5OS44PC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICA8L2N4Om51bURpbT4KICAgICAgICAgICAgPGN4OmV4dExzdD4KICAgICAgICAgICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICAgICAgICAgICAgICA8Y3g6ZXh0IC8+CiAgICAgICAgICAgICAgICA8Y3g6ZXh0IC8+CiAgICAgICAgICAgICAgICA8Y3g6ZXh0IC8+CiAgICAgICAgICAgIDwvY3g6ZXh0THN0PgogICAgICAgIDwvY3g6ZGF0YT4KICAgICAgICA8Y3g6ZGF0YSBpZD0iODM3MiI+CiAgICAgICAgICAgIDxjeDpzdHJEaW0gdHlwZT0iY29sb3JTdHIiPgogICAgICAgICAgICAgICAgPGN4OmY+c3RyaW5nPC9jeDpmPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSI5OTA3IiAvPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSIyMCI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iODY2MiI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSIzMjUwIiBuYW1lPSJzdHJpbmciPgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjU0MTIiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNTIyNSI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIxNjI5Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjYzNDIiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICA8L2N4Omx2bD4KICAgICAgICAgICAgPC9jeDpzdHJEaW0+CiAgICAgICAgICAgIDxjeDpzdHJEaW0gdHlwZT0iZW50aXR5SWQiPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSIxOTIzIj4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI1MzQ5Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjUxMDAiIG5hbWU9InN0cmluZyIgLz4KICAgICAgICAgICAgPC9jeDpzdHJEaW0+CiAgICAgICAgICAgIDxjeDpudW1EaW0gdHlwZT0ieSI+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjU2OTQiIG5hbWU9InN0cmluZyIgLz4KICAgICAgICAgICAgPC9jeDpudW1EaW0+CiAgICAgICAgICAgIDxjeDpleHRMc3Q+CiAgICAgICAgICAgICAgICA8Y3g6ZXh0IC8+CiAgICAgICAgICAgIDwvY3g6ZXh0THN0PgogICAgICAgIDwvY3g6ZGF0YT4KICAgICAgICA8Y3g6ZGF0YSBpZD0iODM5OCI+CiAgICAgICAgICAgIDxjeDpzdHJEaW0gdHlwZT0iZW50aXR5SWQiPgogICAgICAgICAgICAgICAgPGN4OmYgZGlyPSJjb2wiPnN0cmluZzwvY3g6Zj4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iNDIxOCIgbmFtZT0ic3RyaW5nIj4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI5MTI2Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjI5OTEiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNjgwNSI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIzNDI3Ij5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgIDwvY3g6c3RyRGltPgogICAgICAgICAgICA8Y3g6bnVtRGltIHR5cGU9ImNvbG9yVmFsIj4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iMjcwOSIgbmFtZT0ic3RyaW5nIj4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI2MTU1Ij44ODg5MTAwMDAwMDAwLjIzPC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSI5MDgxIiBuYW1lPSJzdHJpbmciIC8+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjEwMjUiIG5hbWU9InN0cmluZyI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iODY0NSI+LTI2MDM2Njk5OTk5OTk5Ljg8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjM3NzUiPjkyNzQ2MDAwMDAwMDAuMjM8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjY2MzUiPi00MTkwMDk5OTk5OTk5OS44PC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI1MzMyIj40OTUyMjEwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSIxNTkyIiBmb3JtYXRDb2RlPSJzdHJpbmciPgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjM2MyI+LTE2NzkxODk5OTk5OTk5Ljg8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjM0MSI+MTcyMzQwMDAwMDAwMC4yMzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNTczNCI+MzU5MjE2MDAwMDAwMDAuMjwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iMjI0MyI+LTM4MzcyOTk5OTk5OTk5Ljg8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjE1NzMiIGZvcm1hdENvZGU9InN0cmluZyIgbmFtZT0ic3RyaW5nIj4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIxNzUyIj4tMzc2Mzc2OTk5OTk5OTkuODwvY3g6cHQ+CiAgICAgICAgICAgICAgICA8L2N4Omx2bD4KICAgICAgICAgICAgPC9jeDpudW1EaW0+CiAgICAgICAgICAgIDxjeDpudW1EaW0gdHlwZT0iY29sb3JWYWwiPgogICAgICAgICAgICAgICAgPGN4OmYgZGlyPSJyb3ciPnN0cmluZzwvY3g6Zj4KICAgICAgICAgICAgICAgIDxjeDpuZiBkaXI9ImNvbCI+c3RyaW5nPC9jeDpuZj4KICAgICAgICAgICAgPC9jeDpudW1EaW0+CiAgICAgICAgICAgIDxjeDpleHRMc3Q+CiAgICAgICAgICAgICAgICA8Y3g6ZXh0IC8+CiAgICAgICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICAgICAgPGN4OmV4dCAvPgogICAgICAgICAgICA8L2N4OmV4dExzdD4KICAgICAgICA8L2N4OmRhdGE+CiAgICAgICAgPGN4OmRhdGEgaWQ9IjQ2NDEiPgogICAgICAgICAgICA8Y3g6bnVtRGltIHR5cGU9ImNvbG9yVmFsIj4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iOTIzNiIgbmFtZT0ic3RyaW5nIj4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIxNjAzIj40NjU5NzIwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI0NDMyIj4tNjk5Njk5OTk5OTk5OS43NjwvY3g6cHQ+CiAgICAgICAgICAgICAgICA8L2N4Omx2bD4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iOTIzNSI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iODY1MyI+LTQ1MjYxOTk5OTk5OTkuNzY8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjY2NjgiPi00MDU2MDg5OTk5OTk5OS44PC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI5OTI3Ij4xODg4NjcwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI1MTUiPjM4MjQyODAwMDAwMDAwLjI8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgIDwvY3g6bnVtRGltPgogICAgICAgICAgICA8Y3g6c3RyRGltIHR5cGU9ImVudGl0eUlkIj4KICAgICAgICAgICAgICAgIDxjeDpmIGRpcj0icm93Ij5zdHJpbmc8L2N4OmY+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjgzMjIiPgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjUzNTgiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNDI1MCI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICA8L2N4OnN0ckRpbT4KICAgICAgICAgICAgPGN4Om51bURpbSB0eXBlPSJzaXplIj4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iODM4OSIgbmFtZT0ic3RyaW5nIj4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI1ODg2Ij4tNDM3Mzg2OTk5OTk5OTkuODwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iMTgyOCI+LTQ1ODgxNzk5OTk5OTk5Ljg8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9Ijg4MzgiPjE1NDMyMzAwMDAwMDAwLjI8L2N4OnB0PgogICAgICAgICAgICAgICAgPC9jeDpsdmw+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjIzNjIiPgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjQ1MyI+NDYxNzYwMDAwMDAwMC4yNDwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNTM0MiI+LTI0MjgyMzk5OTk5OTk5Ljg8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjE3MCI+LTMzMzExMDk5OTk5OTk5Ljg8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9Ijc0OSI+MjI2MDM1MDAwMDAwMDAuMjwvY3g6cHQ+CiAgICAgICAgICAgICAgICA8L2N4Omx2bD4KICAgICAgICAgICAgPC9jeDpudW1EaW0+CiAgICAgICAgICAgIDxjeDpudW1EaW0gdHlwZT0ic2l6ZSI+CiAgICAgICAgICAgICAgICA8Y3g6bHZsIHB0Q291bnQ9IjY4NTQiIGZvcm1hdENvZGU9InN0cmluZyI+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNTkxIj4xNDE4NjIwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI2NTYyIj40MzIzMTYwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIzMjA4Ij4tMzkxOTk5OTk5OTk5OTkuODwvY3g6cHQ+CiAgICAgICAgICAgICAgICA8L2N4Omx2bD4KICAgICAgICAgICAgICAgIDxjeDpsdmwgcHRDb3VudD0iNzkxMCIgbmFtZT0ic3RyaW5nIj4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI3ODM3Ij4tMjA3NDU5OTk5OTk5OTkuODwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iNDAxNyI+MzcwODg2MDAwMDAwMDAuMjwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iMjMwMyI+LTM4MjA5NDk5OTk5OTk5Ljg8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9IjU2NjMiPi0yODAyMzQ5OTk5OTk5OS44PC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSI2NzIwIiAvPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSIyMDkiPgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9Ijk2OTAiPi0yNTYyNjE5OTk5OTk5OS44PC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI5MDYzIj40NjEyNDgwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSIxNTUyIj4zNzM4ODYwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI0MDM1Ij4zODk5OTcwMDAwMDAwMC4yPC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSI2MDE4IiBmb3JtYXRDb2RlPSJzdHJpbmciIG5hbWU9InN0cmluZyIgLz4KICAgICAgICAgICAgPC9jeDpudW1EaW0+CiAgICAgICAgICAgIDxjeDpzdHJEaW0gdHlwZT0iZW50aXR5SWQiPgogICAgICAgICAgICAgICAgPGN4OmY+c3RyaW5nPC9jeDpmPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSI4MSIgbmFtZT0ic3RyaW5nIiAvPgogICAgICAgICAgICAgICAgPGN4Omx2bCBwdENvdW50PSI0OTM2Ij4KICAgICAgICAgICAgICAgICAgICA8Y3g6cHQgaWR4PSI3MTYwIj5zdHJpbmc8L2N4OnB0PgogICAgICAgICAgICAgICAgICAgIDxjeDpwdCBpZHg9Ijk0NzEiPnN0cmluZzwvY3g6cHQ+CiAgICAgICAgICAgICAgICAgICAgPGN4OnB0IGlkeD0iMTA1NCI+c3RyaW5nPC9jeDpwdD4KICAgICAgICAgICAgICAgIDwvY3g6bHZsPgogICAgICAgICAgICA8L2N4OnN0ckRpbT4KICAgICAgICAgICAgPGN4OmV4dExzdCAvPgogICAgICAgIDwvY3g6ZGF0YT4KICAgICAgICA8Y3g6ZXh0THN0PgogICAgICAgICAgICA8Y3g6ZXh0IC8+CiAgICAgICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICAgICAgICAgIDxjeDpleHQgLz4KICAgICAgICA8L2N4OmV4dExzdD4KICAgIDwvY3g6Y2hhcnREYXRhPgogICAgPGN4OmNoYXJ0PgogICAgICAgIDxjeDp0aXRsZSBwb3M9InIiIGFsaWduPSJtaW4iIG92ZXJsYXk9IjEiPgogICAgICAgICAgICA8Y3g6dHg+CiAgICAgICAgICAgICAgICA8Y3g6cmljaD4KICAgICAgICAgICAgICAgICAgICA8Y3g6ZXh0IC8+CiAgICAgICAgICAgICAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgICAgICAgICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICAgICAgICAgIDxjeDpleHQgLz4KICAgICAgICAgICAgICAgIDwvY3g6cmljaD4KICAgICAgICAgICAgPC9jeDp0eD4KICAgICAgICAgICAgPGN4OnNwUHI+CiAgICAgICAgICAgICAgICA8Y3g6ZXh0IC8+CiAgICAgICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICAgICAgPGN4OmV4dCAvPgogICAgICAgICAgICA8L2N4OnNwUHI+CiAgICAgICAgICAgIDxjeDp0eFByPgogICAgICAgICAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgICAgICAgICAgICAgIDxjeDpleHQgLz4KICAgICAgICAgICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICAgICAgICAgIDwvY3g6dHhQcj4KICAgICAgICAgICAgPGN4OmV4dExzdD4KICAgICAgICAgICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICAgICAgICAgIDwvY3g6ZXh0THN0PgogICAgICAgIDwvY3g6dGl0bGU+CiAgICAgICAgPGN4OnBsb3RBcmVhPgogICAgICAgICAgICA8Y3g6cGxvdEFyZWFSZWdpb24+CiAgICAgICAgICAgICAgICA8Y3g6cGxvdFN1cmZhY2UgLz4KICAgICAgICAgICAgICAgIDxjeDpleHRMc3Q+CiAgICAgICAgICAgICAgICAgICAgPGN4OmV4dCAvPgogICAgICAgICAgICAgICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICAgICAgICAgICAgICA8L2N4OmV4dExzdD4KICAgICAgICAgICAgPC9jeDpwbG90QXJlYVJlZ2lvbj4KICAgICAgICAgICAgPGN4OmF4aXMgaWQ9IjM5OTEiIGhpZGRlbj0iMCI+CiAgICAgICAgICAgICAgICA8Y3g6Y2F0U2NhbGluZyBnYXBXaWR0aD0iIiAvPgogICAgICAgICAgICAgICAgPGN4OnRpdGxlPgogICAgICAgICAgICAgICAgICAgIDxjeDp0eD4KICAgICAgICAgICAgICAgICAgICAgICAgPGN4OnJpY2g+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICAgICAgICAgICAgICAgICAgPGN4OmV4dCAvPgogICAgICAgICAgICAgICAgICAgICAgICAgICAgPGN4OmV4dCAvPgogICAgICAgICAgICAgICAgICAgICAgICA8L2N4OnJpY2g+CiAgICAgICAgICAgICAgICAgICAgPC9jeDp0eD4KICAgICAgICAgICAgICAgICAgICA8Y3g6dHhQciAvPgogICAgICAgICAgICAgICAgPC9jeDp0aXRsZT4KICAgICAgICAgICAgICAgIDxjeDp1bml0cyB1bml0PSJodW5kcmVkVGhvdXNhbmRzIj4KICAgICAgICAgICAgICAgICAgICA8Y3g6dW5pdHNMYWJlbD4KICAgICAgICAgICAgICAgICAgICAgICAgPGN4OnR4PgogICAgICAgICAgICAgICAgICAgICAgICAgICAgPGN4OnJpY2ggLz4KICAgICAgICAgICAgICAgICAgICAgICAgPC9jeDp0eD4KICAgICAgICAgICAgICAgICAgICAgICAgPGN4OnR4UHI+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8Y3g6ZXh0IC8+CiAgICAgICAgICAgICAgICAgICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICAgICAgICAgICAgICA8L2N4OnR4UHI+CiAgICAgICAgICAgICAgICAgICAgPC9jeDp1bml0c0xhYmVsPgogICAgICAgICAgICAgICAgPC9jeDp1bml0cz4KICAgICAgICAgICAgICAgIDxjeDptYWpvckdyaWRsaW5lcyAvPgogICAgICAgICAgICAgICAgPGN4OnR4UHIgLz4KICAgICAgICAgICAgICAgIDxjeDpleHRMc3QgLz4KICAgICAgICAgICAgPC9jeDpheGlzPgogICAgICAgICAgICA8Y3g6YXhpcyBpZD0iMTk2MyI+CiAgICAgICAgICAgICAgICA8Y3g6dmFsU2NhbGluZyBtYXg9IjIwODA0MDAwMDAwMDAwLjIiIC8+CiAgICAgICAgICAgICAgICA8Y3g6bWFqb3JHcmlkbGluZXM+CiAgICAgICAgICAgICAgICAgICAgPGN4OmV4dExzdCAvPgogICAgICAgICAgICAgICAgPC9jeDptYWpvckdyaWRsaW5lcz4KICAgICAgICAgICAgICAgIDxjeDptaW5vckdyaWRsaW5lcz4KICAgICAgICAgICAgICAgICAgICA8Y3g6c3BQciAvPgogICAgICAgICAgICAgICAgICAgIDxjeDpleHRMc3Q+CiAgICAgICAgICAgICAgICAgICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICAgICAgICAgICAgICAgICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICAgICAgICAgICAgICAgICAgICAgIDxjeDpleHQgLz4KICAgICAgICAgICAgICAgICAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgICAgICAgICAgICAgICAgICA8L2N4OmV4dExzdD4KICAgICAgICAgICAgICAgIDwvY3g6bWlub3JHcmlkbGluZXM+CiAgICAgICAgICAgICAgICA8Y3g6bWlub3JUaWNrTWFya3M+CiAgICAgICAgICAgICAgICAgICAgPGN4OmV4dExzdD4KICAgICAgICAgICAgICAgICAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgICAgICAgICAgICAgICAgICA8L2N4OmV4dExzdD4KICAgICAgICAgICAgICAgIDwvY3g6bWlub3JUaWNrTWFya3M+CiAgICAgICAgICAgICAgICA8Y3g6dGlja0xhYmVscz4KICAgICAgICAgICAgICAgICAgICA8Y3g6ZXh0THN0PgogICAgICAgICAgICAgICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICAgICAgICAgIDwvY3g6ZXh0THN0PgogICAgICAgICAgICAgICAgPC9jeDp0aWNrTGFiZWxzPgogICAgICAgICAgICAgICAgPGN4Om51bUZtdCBmb3JtYXRDb2RlPSJzdHJpbmciIC8+CiAgICAgICAgICAgICAgICA8Y3g6c3BQcj4KICAgICAgICAgICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICAgICAgICAgIDxjeDpleHQgLz4KICAgICAgICAgICAgICAgICAgICA8Y3g6ZXh0IC8+CiAgICAgICAgICAgICAgICAgICAgPGN4OmV4dCAvPgogICAgICAgICAgICAgICAgPC9jeDpzcFByPgogICAgICAgICAgICAgICAgPGN4OnR4UHIgLz4KICAgICAgICAgICAgPC9jeDpheGlzPgogICAgICAgICAgICA8Y3g6YXhpcyBpZD0iMzkyNiI+CiAgICAgICAgICAgICAgICA8Y3g6Y2F0U2NhbGluZyAvPgogICAgICAgICAgICAgICAgPGN4Om1ham9yR3JpZGxpbmVzPgogICAgICAgICAgICAgICAgICAgIDxjeDpleHRMc3Q+CiAgICAgICAgICAgICAgICAgICAgICAgIDxjeDpleHQgLz4KICAgICAgICAgICAgICAgICAgICAgICAgPGN4OmV4dCAvPgogICAgICAgICAgICAgICAgICAgIDwvY3g6ZXh0THN0PgogICAgICAgICAgICAgICAgPC9jeDptYWpvckdyaWRsaW5lcz4KICAgICAgICAgICAgICAgIDxjeDp0aWNrTGFiZWxzPgogICAgICAgICAgICAgICAgICAgIDxjeDpleHRMc3QgLz4KICAgICAgICAgICAgICAgIDwvY3g6dGlja0xhYmVscz4KICAgICAgICAgICAgPC9jeDpheGlzPgogICAgICAgICAgICA8Y3g6c3BQciAvPgogICAgICAgICAgICA8Y3g6ZXh0THN0PgogICAgICAgICAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgICAgICAgICAgPC9jeDpleHRMc3Q+CiAgICAgICAgPC9jeDpwbG90QXJlYT4KICAgICAgICA8Y3g6bGVnZW5kIHBvcz0ibCIgYWxpZ249ImN0ciIgb3ZlcmxheT0idHJ1ZSI+CiAgICAgICAgICAgIDxjeDpzcFByPgogICAgICAgICAgICAgICAgPGN4OmV4dCAvPgogICAgICAgICAgICAgICAgPGN4OmV4dCAvPgogICAgICAgICAgICA8L2N4OnNwUHI+CiAgICAgICAgICAgIDxjeDp0eFByPgogICAgICAgICAgICAgICAgPGN4OmV4dCAvPgogICAgICAgICAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgICAgICAgICAgICAgIDxjeDpleHQgLz4KICAgICAgICAgICAgPC9jeDp0eFByPgogICAgICAgICAgICA8Y3g6ZXh0THN0PgogICAgICAgICAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgICAgICAgICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICAgICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICA8L2N4OmV4dExzdD4KICAgICAgICA8L2N4OmxlZ2VuZD4KICAgICAgICA8Y3g6ZXh0THN0PgogICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgIDwvY3g6ZXh0THN0PgogICAgPC9jeDpjaGFydD4KICAgIDxjeDpzcFByPgogICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICA8L2N4OnNwUHI+CiAgICA8Y3g6dHhQcj4KICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgPC9jeDp0eFByPgogICAgPGN4OmNsck1hcE92cj4KICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICA8L2N4OmNsck1hcE92cj4KICAgIDxjeDpwcmludFNldHRpbmdzPgogICAgICAgIDxjeDpoZWFkZXJGb290ZXIgYWxpZ25XaXRoTWFyZ2lucz0iMSIgZGlmZmVyZW50T2RkRXZlbj0iMCIgZGlmZmVyZW50Rmlyc3Q9InRydWUiPgogICAgICAgICAgICA8Y3g6b2RkSGVhZGVyPnN0cmluZzwvY3g6b2RkSGVhZGVyPgogICAgICAgICAgICA8Y3g6b2RkRm9vdGVyPnN0cmluZzwvY3g6b2RkRm9vdGVyPgogICAgICAgICAgICA8Y3g6ZXZlbkhlYWRlcj5zdHJpbmc8L2N4OmV2ZW5IZWFkZXI+CiAgICAgICAgICAgIDxjeDpldmVuRm9vdGVyPnN0cmluZzwvY3g6ZXZlbkZvb3Rlcj4KICAgICAgICAgICAgPGN4OmZpcnN0SGVhZGVyPnN0cmluZzwvY3g6Zmlyc3RIZWFkZXI+CiAgICAgICAgICAgIDxjeDpmaXJzdEZvb3Rlcj5zdHJpbmc8L2N4OmZpcnN0Rm9vdGVyPgogICAgICAgIDwvY3g6aGVhZGVyRm9vdGVyPgogICAgICAgIDxjeDpwYWdlTWFyZ2lucyBsPSI2MjM2MzAwMDAwMDAwLjI0IiByPSItMTQ1MDkyOTk5OTk5OTkuOCIgdD0iMTUxOTIwMDAwMDAwMDAuMiIgYj0iMzk3MDgzMDAwMDAwMDAuMiIgaGVhZGVyPSIxNjIwODEwMDAwMDAwMC4yIiBmb290ZXI9IjI4NTAzNDAwMDAwMDAwLjIiIC8+CiAgICAgICAgPGN4OnBhZ2VTZXR1cCBwYXBlclNpemU9IjEiIGZpcnN0UGFnZU51bWJlcj0iMSIgb3JpZW50YXRpb249ImRlZmF1bHQiIGJsYWNrQW5kV2hpdGU9ImZhbHNlIiBkcmFmdD0idHJ1ZSIgdXNlRmlyc3RQYWdlTnVtYmVyPSIxIiBob3Jpem9udGFsRHBpPSI2MDAiIHZlcnRpY2FsRHBpPSI2MDAiIGNvcGllcz0iMSIgLz4KICAgIDwvY3g6cHJpbnRTZXR0aW5ncz4KICAgIDxjeDpleHRMc3Q+CiAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgICAgICA8Y3g6ZXh0IHVyaT0ic3RyaW5nIiAvPgogICAgICAgIDxjeDpleHQgdXJpPSJzdHJpbmciIC8+CiAgICAgICAgPGN4OmV4dCB1cmk9InN0cmluZyIgLz4KICAgIDwvY3g6ZXh0THN0Pgo8L2N4OmNoYXJ0U3BhY2U+";
window['Asc'].document1Xml = '<?xml version="1.0" encoding="utf-8"?>' +
'<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">' +
'\t<cx:chartData>' +
'\t\t<cx:externalData id="string" autoUpdate="1" />' +
'\t\t<cx:data id="2919">' +
'\t\t\t<cx:numDim type="colorVal">' +
'\t\t\t\t<cx:lvl ptCount="8169" formatCode="string">' +
'\t\t\t\t\t<cx:pt idx="5263">43401800000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="6876">4681500000000.24</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="811">-31287599999999.8</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="4533" formatCode="string" name="string">' +
'\t\t\t\t\t<cx:pt idx="304">-11899599999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="3431">45745500000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="5051">21597200000000.2</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:numDim>' +
'\t\t\t<cx:numDim type="size">' +
'\t\t\t\t<cx:f dir="col">string</cx:f>' +
'\t\t\t\t<cx:nf>string</cx:nf>' +
'\t\t\t</cx:numDim>' +
'\t\t\t<cx:strDim type="entityId">' +
'\t\t\t\t<cx:lvl ptCount="8954" name="string">' +
'\t\t\t\t\t<cx:pt idx="1701">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="6441">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="8268">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="2192">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="10000">' +
'\t\t\t\t\t<cx:pt idx="8095">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="1009">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="7291">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:strDim>' +
'\t\t\t<cx:strDim type="cat">' +
'\t\t\t\t<cx:lvl ptCount="6710" name="string">' +
'\t\t\t\t\t<cx:pt idx="7936">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="7645">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="9375">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="2950">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="5228">' +
'\t\t\t\t\t<cx:pt idx="8206">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:strDim>' +
'\t\t\t<cx:extLst>' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t<cx:ext />' +
'\t\t\t\t<cx:ext />' +
'\t\t\t</cx:extLst>' +
'\t\t</cx:data>' +
'\t\t<cx:data id="1070">' +
'\t\t\t<cx:numDim type="y">' +
'\t\t\t\t<cx:lvl ptCount="3609" name="string">' +
'\t\t\t\t\t<cx:pt idx="9586">48881200000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="3871">2274000000000.23</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="2803">657600000000.235</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:numDim>' +
'\t\t\t<cx:strDim type="colorStr">' +
'\t\t\t\t<cx:lvl ptCount="7423" name="string" />' +
'\t\t\t</cx:strDim>' +
'\t\t\t<cx:strDim type="entityId">' +
'\t\t\t\t<cx:f dir="col">string</cx:f>' +
'\t\t\t\t<cx:nf dir="col">string</cx:nf>' +
'\t\t\t\t<cx:lvl ptCount="5681">' +
'\t\t\t\t\t<cx:pt idx="777">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="5130">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="1558">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="8137">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="6425" name="string">' +
'\t\t\t\t\t<cx:pt idx="9945">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="3016">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="645">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="4657" name="string" />' +
'\t\t\t\t<cx:lvl ptCount="2302" name="string">' +
'\t\t\t\t\t<cx:pt idx="198">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="3888">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:strDim>' +
'\t\t\t<cx:strDim type="cat">' +
'\t\t\t\t<cx:lvl ptCount="169">' +
'\t\t\t\t\t<cx:pt idx="5878">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="7722">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="3242">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="4204">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="9749">' +
'\t\t\t\t\t<cx:pt idx="1323">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="8973">' +
'\t\t\t\t\t<cx:pt idx="7850">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="5140">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="3524">' +
'\t\t\t\t\t<cx:pt idx="1212">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="1949">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="4965">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="5813" name="string">' +
'\t\t\t\t\t<cx:pt idx="7841">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:strDim>' +
'\t\t\t<cx:numDim type="size">' +
'\t\t\t\t<cx:f>string</cx:f>' +
'\t\t\t\t<cx:nf>string</cx:nf>' +
'\t\t\t\t<cx:lvl ptCount="2648" formatCode="string" name="string">' +
'\t\t\t\t\t<cx:pt idx="6378">-18028799999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="3690">-25546899999999.8</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:numDim>' +
'\t\t\t<cx:extLst>' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t<cx:ext />' +
'\t\t\t\t<cx:ext />' +
'\t\t\t\t<cx:ext />' +
'\t\t\t</cx:extLst>' +
'\t\t</cx:data>' +
'\t\t<cx:data id="8372">' +
'\t\t\t<cx:strDim type="colorStr">' +
'\t\t\t\t<cx:f>string</cx:f>' +
'\t\t\t\t<cx:lvl ptCount="9907" />' +
'\t\t\t\t<cx:lvl ptCount="20">' +
'\t\t\t\t\t<cx:pt idx="8662">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="3250" name="string">' +
'\t\t\t\t\t<cx:pt idx="5412">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="5225">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="1629">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="6342">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:strDim>' +
'\t\t\t<cx:strDim type="entityId">' +
'\t\t\t\t<cx:lvl ptCount="1923">' +
'\t\t\t\t\t<cx:pt idx="5349">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="5100" name="string" />' +
'\t\t\t</cx:strDim>' +
'\t\t\t<cx:numDim type="y">' +
'\t\t\t\t<cx:lvl ptCount="5694" name="string" />' +
'\t\t\t</cx:numDim>' +
'\t\t\t<cx:extLst>' +
'\t\t\t\t<cx:ext />' +
'\t\t\t</cx:extLst>' +
'\t\t</cx:data>' +
'\t\t<cx:data id="8398">' +
'\t\t\t<cx:strDim type="entityId">' +
'\t\t\t\t<cx:f dir="col">string</cx:f>' +
'\t\t\t\t<cx:lvl ptCount="4218" name="string">' +
'\t\t\t\t\t<cx:pt idx="9126">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="2991">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="6805">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="3427">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:strDim>' +
'\t\t\t<cx:numDim type="colorVal">' +
'\t\t\t\t<cx:lvl ptCount="2709" name="string">' +
'\t\t\t\t\t<cx:pt idx="6155">8889100000000.23</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="9081" name="string" />' +
'\t\t\t\t<cx:lvl ptCount="1025" name="string">' +
'\t\t\t\t\t<cx:pt idx="8645">-26036699999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="3775">9274600000000.23</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="6635">-41900999999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="5332">49522100000000.2</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="1592" formatCode="string">' +
'\t\t\t\t\t<cx:pt idx="363">-16791899999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="341">1723400000000.23</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="5734">35921600000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="2243">-38372999999999.8</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="1573" formatCode="string" name="string">' +
'\t\t\t\t\t<cx:pt idx="1752">-37637699999999.8</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:numDim>' +
'\t\t\t<cx:numDim type="colorVal">' +
'\t\t\t\t<cx:f dir="row">string</cx:f>' +
'\t\t\t\t<cx:nf dir="col">string</cx:nf>' +
'\t\t\t</cx:numDim>' +
'\t\t\t<cx:extLst>' +
'\t\t\t\t<cx:ext />' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t<cx:ext />' +
'\t\t\t</cx:extLst>' +
'\t\t</cx:data>' +
'\t\t<cx:data id="4641">' +
'\t\t\t<cx:numDim type="colorVal">' +
'\t\t\t\t<cx:lvl ptCount="9236" name="string">' +
'\t\t\t\t\t<cx:pt idx="1603">46597200000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="4432">-6996999999999.76</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="9235">' +
'\t\t\t\t\t<cx:pt idx="8653">-4526199999999.76</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="6668">-40560899999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="9927">18886700000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="515">38242800000000.2</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:numDim>' +
'\t\t\t<cx:strDim type="entityId">' +
'\t\t\t\t<cx:f dir="row">string</cx:f>' +
'\t\t\t\t<cx:lvl ptCount="8322">' +
'\t\t\t\t\t<cx:pt idx="5358">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="4250">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:strDim>' +
'\t\t\t<cx:numDim type="size">' +
'\t\t\t\t<cx:lvl ptCount="8389" name="string">' +
'\t\t\t\t\t<cx:pt idx="5886">-43738699999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="1828">-45881799999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="8838">15432300000000.2</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="2362">' +
'\t\t\t\t\t<cx:pt idx="453">4617600000000.24</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="5342">-24282399999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="170">-33311099999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="749">22603500000000.2</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:numDim>' +
'\t\t\t<cx:numDim type="size">' +
'\t\t\t\t<cx:lvl ptCount="6854" formatCode="string">' +
'\t\t\t\t\t<cx:pt idx="591">14186200000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="6562">43231600000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="3208">-39199999999999.8</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="7910" name="string">' +
'\t\t\t\t\t<cx:pt idx="7837">-20745999999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="4017">37088600000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="2303">-38209499999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="5663">-28023499999999.8</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="6720" />' +
'\t\t\t\t<cx:lvl ptCount="209">' +
'\t\t\t\t\t<cx:pt idx="9690">-25626199999999.8</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="9063">46124800000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="1552">37388600000000.2</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="4035">38999700000000.2</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t\t<cx:lvl ptCount="6018" formatCode="string" name="string" />' +
'\t\t\t</cx:numDim>' +
'\t\t\t<cx:strDim type="entityId">' +
'\t\t\t\t<cx:f>string</cx:f>' +
'\t\t\t\t<cx:lvl ptCount="81" name="string" />' +
'\t\t\t\t<cx:lvl ptCount="4936">' +
'\t\t\t\t\t<cx:pt idx="7160">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="9471">string</cx:pt>' +
'\t\t\t\t\t<cx:pt idx="1054">string</cx:pt>' +
'\t\t\t\t</cx:lvl>' +
'\t\t\t</cx:strDim>' +
'\t\t\t<cx:extLst />' +
'\t\t</cx:data>' +
'\t\t<cx:extLst>' +
'\t\t\t<cx:ext />' +
'\t\t\t<cx:ext uri="string" />' +
'\t\t\t<cx:ext />' +
'\t\t</cx:extLst>' +
'\t</cx:chartData>' +
'\t<cx:chart>' +
'\t\t<cx:title pos="r" align="min" overlay="1">' +
'\t\t\t<cx:tx>' +
'\t\t\t\t<cx:rich>' +
'\t\t\t\t\t<cx:ext />' +
'\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t\t<cx:ext />' +
'\t\t\t\t</cx:rich>' +
'\t\t\t</cx:tx>' +
'\t\t\t<cx:spPr>' +
'\t\t\t\t<cx:ext />' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t<cx:ext />' +
'\t\t\t</cx:spPr>' +
'\t\t\t<cx:txPr>' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t<cx:ext />' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t</cx:txPr>' +
'\t\t\t<cx:extLst>' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t</cx:extLst>' +
'\t\t</cx:title>' +
'\t\t<cx:plotArea>' +
'\t\t\t<cx:plotAreaRegion>' +
'\t\t\t\t<cx:plotSurface />' +
'\t\t\t\t<cx:extLst>' +
'\t\t\t\t\t<cx:ext />' +
'\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t</cx:extLst>' +
'\t\t\t</cx:plotAreaRegion>' +
'\t\t\t<cx:axis id="3991" hidden="0">' +
'\t\t\t\t<cx:catScaling gapWidth="" />' +
'\t\t\t\t<cx:title>' +
'\t\t\t\t\t<cx:tx>' +
'\t\t\t\t\t\t<cx:rich>' +
'\t\t\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t\t\t\t<cx:ext />' +
'\t\t\t\t\t\t\t<cx:ext />' +
'\t\t\t\t\t\t</cx:rich>' +
'\t\t\t\t\t</cx:tx>' +
'\t\t\t\t\t<cx:txPr />' +
'\t\t\t\t</cx:title>' +
'\t\t\t\t<cx:units unit="hundredThousands">' +
'\t\t\t\t\t<cx:unitsLabel>' +
'\t\t\t\t\t\t<cx:tx>' +
'\t\t\t\t\t\t\t<cx:rich />' +
'\t\t\t\t\t\t</cx:tx>' +
'\t\t\t\t\t\t<cx:txPr>' +
'\t\t\t\t\t\t\t<cx:ext />' +
'\t\t\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t\t\t</cx:txPr>' +
'\t\t\t\t\t</cx:unitsLabel>' +
'\t\t\t\t</cx:units>' +
'\t\t\t\t<cx:majorGridlines />' +
'\t\t\t\t<cx:txPr />' +
'\t\t\t\t<cx:extLst />' +
'\t\t\t</cx:axis>' +
'\t\t\t<cx:axis id="1963">' +
'\t\t\t\t<cx:valScaling max="20804000000000.2" />' +
'\t\t\t\t<cx:majorGridlines>' +
'\t\t\t\t\t<cx:extLst />' +
'\t\t\t\t</cx:majorGridlines>' +
'\t\t\t\t<cx:minorGridlines>' +
'\t\t\t\t\t<cx:spPr />' +
'\t\t\t\t\t<cx:extLst>' +
'\t\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t\t\t<cx:ext />' +
'\t\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t\t</cx:extLst>' +
'\t\t\t\t</cx:minorGridlines>' +
'\t\t\t\t<cx:minorTickMarks>' +
'\t\t\t\t\t<cx:extLst>' +
'\t\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t\t</cx:extLst>' +
'\t\t\t\t</cx:minorTickMarks>' +
'\t\t\t\t<cx:tickLabels>' +
'\t\t\t\t\t<cx:extLst>' +
'\t\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t\t</cx:extLst>' +
'\t\t\t\t</cx:tickLabels>' +
'\t\t\t\t<cx:numFmt formatCode="string" />' +
'\t\t\t\t<cx:spPr>' +
'\t\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t\t<cx:ext />' +
'\t\t\t\t\t<cx:ext />' +
'\t\t\t\t\t<cx:ext />' +
'\t\t\t\t</cx:spPr>' +
'\t\t\t\t<cx:txPr />' +
'\t\t\t</cx:axis>' +
'\t\t\t<cx:axis id="3926">' +
'\t\t\t\t<cx:catScaling />' +
'\t\t\t\t<cx:majorGridlines>' +
'\t\t\t\t\t<cx:extLst>' +
'\t\t\t\t\t\t<cx:ext />' +
'\t\t\t\t\t\t<cx:ext />' +
'\t\t\t\t\t</cx:extLst>' +
'\t\t\t\t</cx:majorGridlines>' +
'\t\t\t\t<cx:tickLabels>' +
'\t\t\t\t\t<cx:extLst />' +
'\t\t\t\t</cx:tickLabels>' +
'\t\t\t</cx:axis>' +
'\t\t\t<cx:spPr />' +
'\t\t\t<cx:extLst>' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t</cx:extLst>' +
'\t\t</cx:plotArea>' +
'\t\t<cx:legend pos="l" align="ctr" overlay="true">' +
'\t\t\t<cx:spPr>' +
'\t\t\t\t<cx:ext />' +
'\t\t\t\t<cx:ext />' +
'\t\t\t</cx:spPr>' +
'\t\t\t<cx:txPr>' +
'\t\t\t\t<cx:ext />' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t<cx:ext />' +
'\t\t\t</cx:txPr>' +
'\t\t\t<cx:extLst>' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t\t<cx:ext uri="string" />' +
'\t\t\t</cx:extLst>' +
'\t\t</cx:legend>' +
'\t\t<cx:extLst>' +
'\t\t\t<cx:ext uri="string" />' +
'\t\t\t<cx:ext uri="string" />' +
'\t\t\t<cx:ext uri="string" />' +
'\t\t\t<cx:ext uri="string" />' +
'\t\t</cx:extLst>' +
'\t</cx:chart>' +
'\t<cx:spPr>' +
'\t\t<cx:ext uri="string" />' +
'\t\t<cx:ext uri="string" />' +
'\t\t<cx:ext uri="string" />' +
'\t\t<cx:ext uri="string" />' +
'\t</cx:spPr>' +
'\t<cx:txPr>' +
'\t\t<cx:ext uri="string" />' +
'\t</cx:txPr>' +
'\t<cx:clrMapOvr>' +
'\t\t<cx:ext uri="string" />' +
'\t\t<cx:ext uri="string" />' +
'\t</cx:clrMapOvr>' +
'\t<cx:printSettings>' +
'\t\t<cx:headerFooter alignWithMargins="1" differentOddEven="0" differentFirst="true">' +
'\t\t\t<cx:oddHeader>string</cx:oddHeader>' +
'\t\t\t<cx:oddFooter>string</cx:oddFooter>' +
'\t\t\t<cx:evenHeader>string</cx:evenHeader>' +
'\t\t\t<cx:evenFooter>string</cx:evenFooter>' +
'\t\t\t<cx:firstHeader>string</cx:firstHeader>' +
'\t\t\t<cx:firstFooter>string</cx:firstFooter>' +
'\t\t</cx:headerFooter>' +
'\t\t<cx:pageMargins l="6236300000000.24" r="-14509299999999.8" t="15192000000000.2" b="39708300000000.2" header="16208100000000.2" footer="28503400000000.2" />' +
'\t\t<cx:pageSetup paperSize="1" firstPageNumber="1" orientation="default" blackAndWhite="false" draft="true" useFirstPageNumber="1" horizontalDpi="600" verticalDpi="600" copies="1" />' + 
'\t</cx:printSettings>' +
'\t<cx:extLst>' +
'\t\t<cx:ext uri="string" />' +
'\t\t<cx:ext uri="string" />' +
'\t\t<cx:ext uri="string" />' +
'\t\t<cx:ext uri="string" />' +
'\t</cx:extLst>' +
'</cx:chartSpace>';

    window['Asc'].documentXmlGenerated2 = "" +
    "";
