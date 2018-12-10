import base64

try:
    import xlrd
    try:
        from xlrd import xlsx
    except ImportError:
        xlsx = None
except ImportError:
    xlrd = xlsx = None

from odoo import _, api, fields, models
from odoo.exceptions import UserError


class ImportSaleOrderWizard(models.TransientModel):
    _name = 'import.sale.orders.wizard'

    excel_file = fields.Binary(
        placeholder="Choose your excel file.",
        required=True)

    @api.multi
    def action_import_sale_orders(self):
        book = xlrd.open_workbook(
            file_contents=base64.b64decode(self.excel_file))
        sheet = book.sheet_by_index(0)
        col = {
            'client_order_ref': None,
            'partner_id': None,
            'date_order': None,
            'pricelist_id': None,
            'warehouse_id': None,
            'user_id': None,
            'team_id': None,
            'default_code': None,
            'product_uom_qty': None,
        }

        for curr_col in range(0, 9):
            if sheet.cell_value(0, curr_col) in col:
                col[sheet.cell_value(0, curr_col)] = curr_col
        # make sure all columns
        if None in col.values():
            raise UserError(_("Please make sure that "
                              "you have 9 columns in your excel file:\n"
                              "- client_order_ref\n"
                              "- partner_id\n"
                              "- date_order\n"
                              "- pricelist_id\n"
                              "- warehouse_id\n"
                              "- user_id\n"
                              "- team_id\n"
                              "- default_code\n"
                              "- product_uom_qty\n"))
        product_product_env = self.env['product.product']
        sale_order_env = self.env['sale.order']
        res_partner_env = self.env['res.partner']
        res_users_env = self.env['res.users']
        product_pricelist_env = self.env['product.pricelist']
        stock_warehouse_env = self.env['stock.warehouse']
        crm_team_env = self.env['crm.team']
        order_lines = []
        partners = {}  # 'name': id
        pricelists = {}  # 'name': id
        warehouses = {}  # 'name': id
        users = {}  # 'name': id
        teams = {}  # 'name': id
        products = {}  # 'default_code': id

        sheet.cell_value
        current_client_order_ref = sheet.cell_value(1, col['client_order_ref'])
        total_rows = sheet.nrows
        # if client_order_ref changes, create new so
        for curr_row in range(1, total_rows):
            client_order_ref = \
                sheet.cell_value(curr_row, col['client_order_ref'])
            is_last_row = curr_row == (total_rows - 1)
            if current_client_order_ref != client_order_ref or is_last_row:
                if is_last_row:
                    row = curr_row
                    products, order_lines = \
                        self.ensure_order_line(sheet, row, col, product_product_env, products, order_lines)
                else:
                    row = curr_row - 1
                # convert to correct time format
                date_order = \
                    fields.datetime.strptime(
                        sheet.cell_value(
                            row, col['date_order']), '%m/%d/%Y'
                    ).strftime('%Y-%m-%d %H:%M:%S')
                # ensure partner_id
                partner_name = sheet.cell_value(row, col['partner_id'])
                if partner_name not in partners:
                    partners[partner_name] = self.ensure_id(res_partner_env, partner_name)
                # ensure pricelist_id
                pricelist_name = sheet.cell_value(row, col['pricelist_id'])
                if pricelist_name not in pricelists:
                    pricelists[pricelist_name] = self.ensure_id(product_pricelist_env, pricelist_name)
                # ensure warehouse_id
                warehouses_name = sheet.cell_value(row, col['warehouse_id'])
                if warehouses_name not in warehouses:
                    warehouses[warehouses_name] = self.ensure_id(stock_warehouse_env, warehouses_name)
                # ensure user_id
                user_name = sheet.cell_value(row, col['user_id'])
                if user_name not in users:
                    users[user_name] = self.ensure_id(res_users_env, user_name)
                # ensure team_id
                team_name = sheet.cell_value(row, col['team_id'])
                if team_name not in teams:
                    teams[team_name] = self.ensure_id(crm_team_env, team_name)
                sale_order_env.create({
                    'client_order_ref': current_client_order_ref,
                    'partner_id': partners[partner_name],
                    'date_order': date_order,
                    'pricelist_id': pricelists[pricelist_name],
                    'warehouse_id': warehouses[warehouses_name],
                    'user_id': users[user_name],
                    'team_id': teams[team_name],
                    'order_line': order_lines
                })
                current_client_order_ref = client_order_ref
                order_lines = []  # reset order lines
            # append new order_line into order_lines
            # ensure product_id
            products, order_lines = \
                self.ensure_order_line(sheet, curr_row, col, product_product_env, products, order_lines)

    @api.model
    def ensure_order_line(self, sheet, curr_row, col, product_product_env, products, order_lines):
        product_code = sheet.cell_value(curr_row, col['default_code'])
        if product_code not in products:
            products[product_code] = \
                self.ensure_product_id(product_product_env, product_code)
        order_lines.append((0, 0, {
            'product_id': products[product_code],
            'product_uom_qty':
                sheet.cell_value(curr_row, col['product_uom_qty']),
        }))
        return products, order_lines

    @api.model
    def ensure_id(self, env, name):
        obj_id = env.search([('name', '=', name)], limit=1).id
        if obj_id:
            return obj_id
        raise UserError(_("%s is not exist" % name))

    @api.model
    def ensure_product_id(
            self, env, default_code):
        if default_code:
            product = env.search([
                ('default_code', '=', default_code)
            ], limit=1)
            if product:
                return product.id
        raise UserError(_(
            "default_code: %s is not exist" % default_code))
