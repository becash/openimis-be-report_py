import datetime
import io
import json
import logging
import os
import tempfile

import openpyxl
import openpyxl.styles
import openpyxl.utils
from django.http import Http404, HttpResponse, HttpResponseBadRequest, FileResponse
from django.template import loader
from django.utils.translation import gettext as _
from django.views.decorators.clickjacking import xframe_options_exempt
from django.views.decorators.http import require_http_methods
from im_export.views import check_user_rights
from insuree.apps import InsureeConfig
from reportbro import Report, ReportBroError
from rest_framework.decorators import api_view, permission_classes
from rest_framework.exceptions import PermissionDenied

from report.services import generate_report, get_report_definition
from worker_voucher.models import WorkerVoucher
from worker_voucher.services import get_voucher_user_filters

from .apps import ReportConfig

logger = logging.getLogger(__file__)


@api_view(["GET"])
def report(request, report_name, report_format="pdf", alternate=None):
    """
    Run a report
    :param request: Predefined by Django
    :param report_name: Report name within the module
    :param report_format: pdf (default) or xlsx
    :param alternate: Future use, allows several templates for a single report: different languages or report variants
    :return: view
    """
    logger.debug(
        "report name %s in %s format",
        report_name,
        report_format,
    )
    report_config = ReportConfig.get_report(report_name)
    if not report_config:
        raise Http404("Poll does not exist")
    report_definition = get_report_definition(
        report_name, report_config["default_report"]
    )
    if (
        report_config.get("permission")
        and not request.user.has_perms(ReportConfig.gql_query_report_perms)
        and not request.user.has_perms(report_config.get("permission"))
    ):
        raise PermissionDenied(_("unauthorized"))

    # parameters tend to get put in lists because they *could* be repeated
    unlisted = {
        k: (v[0] if isinstance(v, list) and len(v) == 0 else v)
        for k, v in request.GET.items()
    }

    data = report_config["python_query"](request.user, **unlisted)

    return FileResponse(
        io.BytesIO(
            generate_report(
                report_name,
                report_definition,
                data,
                report_format,
            )
        ), filename=f"{report_name}.{report_format}", as_attachment=False
    )


@xframe_options_exempt
def reportbro_designer(request):
    template = loader.get_template("report/reportbro.html")

    context = {}
    return HttpResponse(template.render(context, request))


@xframe_options_exempt
def reportbro_previewer(request):
    """
    Generates a report preview within the designer. This can work in two ways:
    1. The report details are passed as a PUT. We generate the report and store the result in a temporary file.
       The Designer then runs a GET request to retrieve the generated report with the key returned by the PUT request.
       This only generates PDFs in theory.
    2. The report is generated on the fly. This is used by the Designer to generate PDFs and XLSX files.
    """
    response = HttpResponse('')
    response['Access-Control-Allow-Origin'] = '*'
    response['Access-Control-Allow-Methods'] = 'GET, PUT, OPTIONS'
    response['Access-Control-Allow-Headers'] = \
        'Origin, X-Requested-With, X-HTTP-Method-Override, Content-Type, Accept, Authorization, Z-Key'

    if request.method == "PUT":
        json_data = json.loads(request.body.decode('utf-8'))
        output_format = json_data.get('outputFormat')
        if output_format not in ('pdf', 'xlsx'):
            return HttpResponseBadRequest('outputFormat parameter missing or invalid')
        if not isinstance(json_data, dict) or not isinstance(json_data.get('report'), dict) or \
                not isinstance(json_data.get('data'), dict) or not isinstance(json_data.get('isTestData'), bool):
            return HttpResponseBadRequest('invalid report values')
        report_definition = json_data.get('report')
        data = json_data.get('data')
        is_test_data = json_data.get('isTestData')
        try:
            temp_file=tempfile.NamedTemporaryFile(delete=False)
            key = os.path.basename(temp_file.name)
            generate_report("preview", report_definition, data, output_format,
                            local_file=temp_file.name, is_test_data=is_test_data)
            return HttpResponse('key:'+key)
        except ReportBroError as e:
            logger.exception(e.error)
            return HttpResponse(json.dumps(dict(errors=[e.error])))
        except Exception as e:
            logger.exception(e)
            return HttpResponseBadRequest('failed to generate report: ' + str(e))
    if request.method == 'GET':
        output_format = request.GET.get('outputFormat')
        if output_format not in ('pdf', 'xlsx'):
            return HttpResponseBadRequest('outputFormat parameter missing or invalid')
        key = request.GET.get('key')
        if not key:
            # in case there is a GET request without a key we expect all report data to be available.
            # this is NOT used by ReportBro Designer and only added for the sake of completeness.
            json_data = json.loads(request.body.decode('utf-8'))
            if not isinstance(json_data, dict) or not isinstance(json_data.get('report'), dict) or \
                    not isinstance(json_data.get('data'), dict) or not isinstance(json_data.get('isTestData'), bool):
                return HttpResponseBadRequest('invalid report values')
            report_definition = json_data.get('report')
            data = json_data.get('data')
            is_test_data = json_data.get('isTestData')
            if not isinstance(report_definition, dict) or not isinstance(data, dict):
                return HttpResponseBadRequest('report_definition or data missing')
            return FileResponse(
                io.BytesIO(
                    generate_report(
                        "preview",
                        report_definition,
                        data,
                        output_format,
                        is_test_data=is_test_data,
                    )
                ), filename=f"preview.{output_format}", as_attachment=False
            )
        try:
            with open(os.path.join(tempfile.gettempdir(), key), 'rb') as f:
                response = HttpResponse(f.read(), content_type='application/pdf')
                response['Content-Disposition'] = 'inline; filename="report_preview.pdf"'
                return response
        except Exception as e:
            logger.exception(e)
            return HttpResponseBadRequest('failed to generate report: ' + str(e))
    return HttpResponseBadRequest('invalid request method')


@require_http_methods(["GET"])
def download_worker_vouchers_xlsx(request):
    user = request.user

    if user.is_anonymous:
        return HttpResponseBadRequest('User is not authenticated')

    # Get query parameters
    date_from = request.GET.get('date_from')
    date_to = request.GET.get('date_to')

    if not date_from or not date_to:
        return HttpResponseBadRequest('Missing required parameters: date_from,date_to')

    try:
        if date_from:
            date_from = datetime.datetime.strptime(date_from, '%Y-%m-%d')
        if date_to:
            date_to = datetime.datetime.strptime(date_to, '%Y-%m-%d')
    except ValueError:
        return HttpResponse('Invalid date format. Use YYYY-MM-DD', status=400)

    # Build query
    queryset = WorkerVoucher.objects.all()

    queryset = queryset.filter(policyholder_id=user.id)

    # Apply filters
    # if status:
    #     if status not in [choice[0] for choice in WorkerVoucher.Status.choices]:
    #         return HttpResponse('Invalid status', status=400)
    #     queryset = queryset.filter(status=status)

    if date_from:
        queryset = queryset.filter(assigned_date__gte=date_from)
    if date_to:
        queryset = queryset.filter(assigned_date__lte=date_to)


    # eu = PolicyHolder.objects.filter(economic_unit_user_filter(info.context.user), code=economic_unit_code).first()
    # if not eu:
    #     raise AttributeError(_("workers.validation.economic_unit_not_exist"))
    #
    # query = Insuree.get_queryset(None, info.context.user).distinct('id').filter(
    #     worker_user_filter(info.context.user, economic_unit_code=economic_unit_code),
    #     workervoucher__is_deleted=False,
    #     workervoucher__policyholder__is_deleted=False,
    #     workervoucher__policyholder__code=economic_unit_code,
    # )

    # Order by assigned date
    queryset = queryset.order_by('assigned_date')
    # Build query
    queryset = WorkerVoucher.objects.all()

    # Apply filters
    queryset = queryset.filter(status=WorkerVoucher.Status.ASSIGNED)

    if date_from:
        queryset = queryset.filter(assigned_date__gte=date_from)
    if date_to:
        queryset = queryset.filter(assigned_date__lte=date_to)

    # Get user-specific filters
    # queryset = WorkerVoucher.get_queryset(queryset, request.user)

    # Order by assigned date
    queryset = queryset.order_by('assigned_date')

    wb = create_download_worker_vouchers_xlsx(queryset)

    # Create response with Excel file
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="Report {date_from} - {date_to}.xlsx"'

    # Save to response
    wb.save(response)
    return response


def create_download_worker_vouchers_xlsx(worker_vouchers_queryset):
    wb = openpyxl.Workbook()
    ws = wb.active

    # Define styles
    border = openpyxl.styles.Border(
        left=openpyxl.styles.Side(style='thin'),
        right=openpyxl.styles.Side(style='thin'),
        top=openpyxl.styles.Side(style='thin'),
        bottom=openpyxl.styles.Side(style='thin')
    )

    center_aligned = openpyxl.styles.Alignment(horizontal='center',
                               vertical='center',
                               wrap_text=True)

    # Headers (first row with merged cells)
    headers_row1 = [
        'Nr.\nd/o',  # A
        'Cod voucher',  # B
        'Data desfăşurării activităţii',  # B
        'Ora',  # C-D (merged)
        '',  # D (part of merged Ora)
        'Numele şi prenumele zilierului',  # E
        'IDNP, domi-ciliu',  # F
        'Semnătura zilierului la începerea activităţii',  # G
        'Semnătura zilierului privind instruirea',  # H
        'Locul exercitării activităţii',  # I
        'Denumi-rea activităţii desfăşurate',  # J
        'Remune-rația negociată, în cifre şi în litere, lei',  # K
        'Remunerația achitată, în cifre şi în litere, lei',  # L
        'Semnătura de confirmare a primirii banilor',  # M
        'Locul pentru ştampilă şi semnătura beneficiarului'  # N
    ]

    # Second row headers (under Ora)
    headers_row2 = [
        '',  # A (merged with above)
        '',  # B (merged with above)
        '',  # B (merged with above)
        'începerii zilei de lucru',  # C
        'finalizării zilei de lucru',  # D
        '',  # E
        '',  # F
        '',  # G
        '',  # H
        '',  # I
        '',  # J
        '',  # K
        '',  # L
        '',  # M
        ''  # N
    ]

    # Write headers and apply styles
    for col, header in enumerate(headers_row1, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.border = border
        cell.alignment = center_aligned
        cell.font = openpyxl.styles.Font(name='Times New Roman', size=12)

    # Merge 'Ora' cells
    ws.merge_cells('D1:E1')

    # Write second row and apply styles
    for col, header in enumerate(headers_row2, 1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.border = border
        cell.alignment = center_aligned
        cell.font = openpyxl.styles.Font(name='Times New Roman', size=12)

    # Merge cells for first two columns
    for letter in ["A","B","C","F","G","H","I","J","K","L","M","N","O"]:
        ws.merge_cells(letter+'1:'+letter+'2')

    # Add numbers row
    numbers = range(1, 16)
    for col, num in enumerate(numbers, 1):
        cell = ws.cell(row=3, column=col, value=num)
        cell.border = border
        cell.alignment = center_aligned
        cell.font = openpyxl.styles.Font(name='Times New Roman', size=14)

    # Set column widths
    for col in range(1, 15):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15

    # Set row heights
    ws.row_dimensions[1].height = 100  # First row higher for wrapped text
    ws.row_dimensions[2].height = 40  # Second row
    ws.row_dimensions[3].height = 25  # Numbers row

    # Populate data
    row = 4  # Start after headers and number row
    for index, voucher in enumerate(worker_vouchers_queryset, 1):
        ws.cell(row=row, column=1, value=voucher.insuree.id)  # Nr. d/o
        ws.cell(row=row, column=2, value=str(voucher.id))
        ws.cell(row=row, column=3, value=voucher.assigned_date.strftime('%Y-%m-%d') if voucher.assigned_date else '')
        ws.cell(row=row, column=4, value=voucher.start_time.strftime('%H:%M') if voucher.start_time else '')
        ws.cell(row=row, column=5, value=voucher.end_time.strftime('%H:%M') if voucher.end_time else '')
        ws.cell(row=row, column=6,
                value=f"{voucher.insuree.other_names} {voucher.insuree.last_name}" if voucher.insuree else '')
        ws.cell(row=row, column=7, value=voucher.insuree.chf_id if voucher.insuree else '')
        # Columns 7 and 8 are for signatures, leave them empty
        ws.cell(row=row, column=10, value=voucher.work_place or '')
        ws.cell(row=row, column=11, value=voucher.activity or '')
        ws.cell(row=row, column=12, value=str(voucher.negotiated) if voucher.negotiated else '')
        ws.cell(row=row, column=13, value=str(voucher.paid) if voucher.paid else '')
        # Columns 13 and 14 are for signatures and stamps, leave them empty

        # Apply styles to all cells in the row
        for col in range(1, 16):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = center_aligned

        row += 1

    return wb
