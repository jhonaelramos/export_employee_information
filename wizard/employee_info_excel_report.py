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
from odoo import api, fields, models, _
import xlwt, xlsxwriter
import base64


class EmployeeInfoExcelReport(models.TransientModel):
    _name = 'employee.info.excel.report'
    _description = 'Exportar a Excel la Data de Luchadores'

    company_id = fields.Many2one('res.company', string='Company', required=True)

    @api.multi
    def generated_excel_report(self, record):
        employee_obj = self.env['hr.employee']
        employee_search = employee_obj.search([('company_id', '=', self.company_id.id)])
        workbook = xlwt.Workbook()

        # Style for Excel Report
        style0 = xlwt.easyxf('align: horiz left; borders: ; pattern: pattern solid, fore_colour white;', num_format_str='#0')
        style1 = xlwt.easyxf('font:bold True, color Yellow , height 480;  borders:top double; align: horiz center; pattern: pattern solid, fore_colour gray25;', num_format_str='#,##0.00')
        style2 = xlwt.easyxf('font:bold True, color white, height 480;  borders:top double; align: horiz center; pattern: pattern solid, fore_colour red;', num_format_str='#,##0.00')
        styletitle = xlwt.easyxf(
            'font:bold True, color white, height 400;  borders: top double; align: horiz center; pattern: pattern solid, fore_colour gray25;',
            num_format_str='#,##0.00')
        sheet = workbook.add_sheet("Employee Information List")

        sheet.write_merge(0, 0, 0, 12, 'Informacion del Luchador', style2)

        sheet.write(1, 0, 'Estado', styletitle)
        sheet.write(1, 1, 'Municipio', styletitle)
        sheet.write(1, 2, 'Parroquia', styletitle)
        sheet.write(1, 3, 'Cedula', styletitle)
        sheet.write(1, 4, 'Nombre', styletitle)
        sheet.write(1, 5, 'Telefono', styletitle)
        sheet.write(1, 6, 'Movil', styletitle)
        sheet.write(1, 7, 'Edad', styletitle)
        sheet.write(1, 8, 'Avanzada', styletitle)
        sheet.write(1, 9, 'Sexo', styletitle)
        sheet.write(1, 10, 'Responsabilidad', styletitle)
        sheet.write(1, 11, 'Serial', styletitle)

        sheet.col(0).width = 700 * (len('Estado') + 1)
        sheet.col(1).width = 700 * (len('Municipio') + 1)
        sheet.col(2).width = 700 * (len('Parroquia') + 1)
        sheet.col(3).width = 700 * (len('Cedula') + 1)
        sheet.col(4).width = 700 * (len('Nombre') + 1)
        sheet.col(5).width = 700 * (len('Telefono') + 1)
        sheet.col(6).width = 700 * (len('Movil') + 1)
        sheet.col(7).width = 700 * (len('Edad') + 1)
        sheet.col(8).width = 700 * (len('Avanzada') + 1)
        sheet.col(9).width = 700 * (len('Sexo') + 1)
        sheet.col(10).width = 700 * (len('Responsabilidad') + 1)
        sheet.col(11).width = 700 * (len('Serial') + 1)
        sheet.row(0).height_mismatch = True
        sheet.row(0).height = 256 * 2
        sheet.row(1).height = 256 * 2
        sheet.row(2).height = 256 * 2

        row = 2
        for rec in employee_search:
            sheet.write(row, 0, rec.state_id.name, style0)
            sheet.write(row, 1, rec.municipality_id.name, style0)
            sheet.write(row, 2, rec.parish_id.name, style0)
            sheet.write(row, 3, rec.identification_id, style0)
            sheet.write(row, 4, rec.name, style0)
            sheet.write(row, 5, rec.mobile_phone, style0)
            sheet.write(row, 6, rec.work_phone, style0)
            sheet.write(row, 7, rec.edad, style0)
            sheet.write(row, 8, rec.avanzada_id.name, style0)
            sheet.write(row, 9, rec.gender, style0)
            sheet.write(row, 10, rec.responsabilidad_ffm_id.name, style0)
            sheet.write(row, 11, rec.serial_ciudadano, style0)
            row +=1
        workbook.save('/tmp/employee_info_list.xls')
        result_file = open('/tmp/employee_info_list.xls', 'rb').read()
        attachment_id = self.env['wizard.emp.info.excel.report'].create({
            'name': 'Data del luchador.xls',
            'report': base64.encodestring(result_file)
        })

        return {
            'name': _('Notification'),
            'context': self.env.context,
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'wizard.emp.info.excel.report',
            'res_id': attachment_id.id,
            'data': None,
            'type': 'ir.actions.act_window',
            'target': 'new'
        }


class NucleoInfoExcelReport(models.TransientModel):
    _name = 'nucleo.info.excel.report'
    _description = 'Exportar a Excel la Data de Nucleo'

    company_id = fields.Many2one('res.company', string='Company', required=True)

    @api.multi
    def generated_excel_report2(self, record):
        employee_obj = self.env['control.centro']
        # employee_search = employee_obj.search([('company_id', '=', self.company_id.id)])
        workbook = xlwt.Workbook()

        # Style for Excel Report
        style0 = xlwt.easyxf('align: horiz left; borders: ; pattern: pattern solid, fore_colour white;', num_format_str='#0')
        style1 = xlwt.easyxf('font:bold True, color Yellow , height 480;  borders:top double; align: horiz center; pattern: pattern solid, fore_colour gray25;', num_format_str='#,##0.00')
        style2 = xlwt.easyxf('font:bold True, color white, height 480;  borders:top double; align: horiz center; pattern: pattern solid, fore_colour red;', num_format_str='#,##0.00')
        styletitle = xlwt.easyxf(
            'font:bold True, color white, height 400;  borders: top double; align: horiz center; pattern: pattern solid, fore_colour gray25;',
            num_format_str='#,##0.00')
        sheet = workbook.add_sheet("Employee Information List")

        sheet.write_merge(0, 0, 0, 10, 'Informacion del Luchador', style2)

        sheet.write(1, 0, 'Estado', styletitle)
        sheet.write(1, 1, 'Municipio', styletitle)
        sheet.write(1, 2, 'Parroquia', styletitle)
        sheet.write(1, 3, 'Centro', styletitle)
        sheet.write(1, 4, 'Codigo', styletitle)
        sheet.write(1, 5, 'Cedula Jefe de Nucleo LSB', styletitle)
        sheet.write(1, 6, 'Nombre Jefe de Nucleo LSB', styletitle)
        sheet.write(1, 7, 'Telefono Jefe de Nucleo LSB', styletitle)
        sheet.write(1, 8, 'Cedula Formador de Nucleo LSB', styletitle)
        sheet.write(1, 9, 'Nombre Formador de Nucleo LSB', styletitle)
        sheet.write(1, 10, 'Telefono Formador de Nucleo LSB', styletitle)
        sheet.write(1, 11, 'Cedula Agitador de Nucleo LSB', styletitle)
        sheet.write(1, 12, 'Nombre Agitador de Nucleo LSB', styletitle)
        sheet.write(1, 13, 'Telefono Agitador de Nucleo LSB', styletitle)
        sheet.write(1, 14, 'Cedula Organizador de Nucleo LSB', styletitle)
        sheet.write(1, 15, 'Nombre Organizador de Nucleo LSB', styletitle)
        sheet.write(1, 16, 'Telefono Organizador de Nucleo LSB', styletitle)

        sheet.col(0).width = 700 * (len('Estado') + 1)
        sheet.col(1).width = 700 * (len('Municipio') + 1)
        sheet.col(2).width = 700 * (len('Parroquia') + 1)
        sheet.col(3).width = 700 * (len('Centro') + 1)
        sheet.col(4).width = 700 * (len('Codigo') + 1)
        sheet.col(5).width = 700 * (len('Cedula Jefe de Nucleo LSB') + 1)
        sheet.col(6).width = 700 * (len('Nombre Jefe de Nucleo LSB') + 1)
        sheet.col(7).width = 700 * (len('Telefono Jefe de Nucleo LSB') + 1)
        sheet.col(8).width = 700 * (len('Cedula Formador de Nucleo LSB') + 1)
        sheet.col(9).width = 700 * (len('Nombre Formador de Nucleo LSB') + 1)
        sheet.col(10).width = 700 * (len('Telefono Formador de Nucleo LSB') + 1)
        sheet.col(11).width = 700 * (len('Cedula Agitador de Nucleo LSB') + 1)
        sheet.col(12).width = 700 * (len('Nombre Agitador de Nucleo LSB') + 1)
        sheet.col(13).width = 700 * (len('Telefono Agitador de Nucleo LSB') + 1)
        sheet.col(14).width = 700 * (len('Cedula Organizador de Nucleo LSB') + 1)
        sheet.col(15).width = 700 * (len('Nombre Organizador de Nucleo LSB') + 1)
        sheet.col(16).width = 700 * (len('Telefono Organizador de Nucleo LSB') + 1)
        sheet.row(0).height_mismatch = True
        sheet.row(0).height = 256 * 2
        sheet.row(1).height = 256 * 2
        sheet.row(2).height = 256 * 2

        row = 2
        for rec in employee_obj:
            sheet.write(row, 0, rec.state_id.name, style0)
            sheet.write(row, 1, rec.municipality_id.name, style0)
            sheet.write(row, 2, rec.parish_id.name, style0)
            sheet.write(row, 3, rec.centro.name, style0)
            sheet.write(row, 4, rec.centro.codigo, style0)
            sheet.write(row, 5, rec.responsable_id.identification_id, style0)
            sheet.write(row, 6, rec.responsable_id.name, style0)
            sheet.write(row, 7, rec.responsable_id.mobile_phone, style0)
            sheet.write(row, 8, rec.formador_id.identification_id, style0)
            sheet.write(row, 9, rec.formador_id.name, style0)
            sheet.write(row, 10, rec.formador_id.mobile_phone, style0)
            sheet.write(row, 11, rec.agitador_id.identification_id, style0)
            sheet.write(row, 12, rec.agitador_id.name, style0)
            sheet.write(row, 13, rec.agitador_id.mobile_phone, style0)
            sheet.write(row, 14, rec.organizador_id.identification_id, style0)
            sheet.write(row, 15, rec.organizador_id.name, style0)
            sheet.write(row, 16, rec.organizador_id.mobile_phone, style0)
            row +=1
        workbook.save('/tmp/nucleo_info_list.xls')
        result_file = open('/tmp/nucleo_info_list.xls', 'rb').read()
        attachment_id = self.env['wizard.nucleo.info.excel.report'].create({
            'name': 'Data del Nucleo.xls',
            'report': base64.encodestring(result_file)
        })

        return {
            'name': _('Notification'),
            'context': self.env.context,
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'wizard.nucleo.info.excel.report',
            'res_id': attachment_id.id,
            'data': None,
            'type': 'ir.actions.act_window',
            'target': 'new'
        }

class WizardEmployeeInformationExcelReport(models.TransientModel):
    _name = 'wizard.emp.info.excel.report'

    name = fields.Char('File Name', size=64)
    report = fields.Binary('Prepared File', filters='.xls', readonly=True)

class WizardNucleoInformationExcelReport(models.TransientModel):
    _name = 'wizard.nucleo.info.excel.report'

    name = fields.Char('File Name', size=64)
    report = fields.Binary('Prepared File', filters='.xls', readonly=True)	