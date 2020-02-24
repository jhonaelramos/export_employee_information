# -*- coding: utf-8 -*-
##############################################################################
#
#    OdooDevelopers.
#    Copyright (C) 2017-TODAY OdooDevelopers(<http://www.odoodevelopers.com>).
#    Author: Redouane ADADI(<http://www.odoodevelopers.com>)
#    you can modify it under the terms of the GNU LESSER
#    GENERAL PUBLIC LICENSE (LGPL v3), Version 3.
#
#    It is forbidden to publish, distribute, sublicense, or sell copies
#    of the Software or modified copies of the Software.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU LESSER GENERAL PUBLIC LICENSE (LGPL v3) for more details.
#
#    You should have received a copy of the GNU LESSER GENERAL PUBLIC LICENSE
#    GENERAL PUBLIC LICENSE (LGPL v3) along with this program.
#    If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################
{
    'name': 'Employee Information Excel Report',
    'version': '1.1',
    'summary': 'Employee Information Excel Report',
    'description': 'Employee Information Excel Report',
    'category': 'Human Resources',
    'author': 'OdooDevelopers',
    'website': 'http://www.odoodevelopers.com',
    'company': 'Odoo Developers',
    'depends': ['base', 'hr', 'report_xlsx'],
    'data': ['wizard/employee_info_excel_report.xml'],
    'images': ['static/description/main.png'],
    'license': 'AGPL-3',
    'installable': True,
    'auto_install': False,
    'application': False,
}