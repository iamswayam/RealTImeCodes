@login_required
@program_required
def admin_benefit_summary(request):
    program = request.user.get_current_program()
    role_users_list = program.get_program_roles_as_choices()
    form = forms.AdminBenefitsSummaryReportForm(role_users_list, request.POST or None)
    ctx = {}
    if request.method == "POST":
        headers = [
            "Last Name",
            "First Name",
            "Role",
            "Primary Program",
            "Home Institution",
        ]
        category_names = list(
            BudgetCategory.objects.all().only("name").values_list("name", flat=True)
        )
        headers.extend(category_names)
        headers.append("Total (All Time Off Accounts)")

        xlsx_data = ()
        xlsx_data += (headers,)

        from_date = request.POST.get("from_date", None)
        to_date = request.POST.get("to_date", None)
        roles = request.POST.getlist("roles", [])
        start = make_aware(datetime.strptime(from_date, "%m/%d/%Y"))
        end = make_aware(datetime.strptime(to_date, "%m/%d/%Y"))

        queryset = (
            Account.objects.filter(
                Q(users__profile__roles__role__name__in=roles)
                & Q(users__profile__roles__active=True)
                & Q(users__profile__affiliation__program=program)
                & Q(transaction_history__executed_date__date__gte=start)
                & Q(transaction_history__executed_date__date__lte=end)
            )
            .prefetch_related(
                Prefetch(
                    "transaction_history",
                    queryset=AccountTransaction.objects.filter(
                        Q(account__parent__isnull=False)
                        & Q(request__specific_date__date__gte=start)
                        & Q(request__end_date__date__lte=end)
                    ),
                    to_attr="my_debits",
                ),
                Prefetch(
                    "users",
                    queryset=User.objects.filter(profile__roles__active=True)
                    .annotate(
                        user_role=F("profile__roles__role__name"),
                    )
                    .prefetch_related(
                        Prefetch(
                            "profile",
                            queryset=UserProfile.objects.all().prefetch_related(
                                Prefetch(
                                    "affiliation_set",
                                    queryset=Affiliation.objects.filter(
                                        primary_program=True
                                    )
                                    .prefetch_related("program")
                                    .all(),
                                    to_attr="primary_program_data",
                                ),
                                Prefetch(
                                    "gmeprogramhistory_set",
                                    queryset=GMEProgramHistory.objects.filter(
                                        site__isnull=False
                                    )
                                    .select_related("site")
                                    .all(),
                                    to_attr="home_institution_data",
                                ),
                            ),
                            to_attr="profile_data",
                        ),
                    ),
                    to_attr="users_data",
                ),
            )
            .order_by("users")
        )

        output = {}
        for index, each in enumerate(queryset):
            user = each.users_data[0]
            if index == 0:
                home_institution = None
                primary_program = None
                try:
                    primary_program = user.profile_data.primary_program_data[0]
                except Exception:
                    primary_program = None

                try:
                    home_institution = user.profile_data.home_institution_data[0]
                except Exception:
                    home_institution = None

                output[user] = {
                    "first_name": user.first_name,
                    "last_name": user.last_name,
                    "role": user.user_role,
                    "primary_program": primary_program.program.name
                    if primary_program
                    else "",
                    "home_institution": home_institution.gme_program
                    if home_institution
                    else "",
                }
                # Initialises counts for time off categories
                for category_name in category_names:
                    output[user][category_name] = 0

            if user in output:
                leaves_used = 0
                for each_debit in each.my_debits:
                    if each_debit.debit:
                        leaves_used += each_debit.value
                    else:
                        leaves_used -= each_debit.value
                output[user][each.name] = round(abs(leaves_used) / 1440)

            else:
                home_institution = None
                primary_program = None
                try:
                    primary_program = user.profile_data.primary_program_data[0]
                except Exception:
                    primary_program = None

                try:
                    home_institution = user.profile_data.home_institution_data[0]
                except Exception:
                    home_institution = None

                output[user] = {
                    "first_name": user.first_name,
                    "last_name": user.last_name,
                    "role": user.user_role,
                    "primary_program": primary_program.program.name
                    if primary_program
                    else "",
                    "home_institution": home_institution.gme_program
                    if home_institution
                    else "",
                }
                # Initialises counts for time off categories
                for category_name in category_names:
                    output[user][category_name] = 0

        for user, data in output.items():
            list_of_values = list(data.values())
            count_value = list_of_values[
                -len(category_names) :
            ]  # Defining "count_value" to get last 'n' elements of value from "list_of_values"
            data["Total"] = sum(count_value)

            xlsx_data += (list(data.values()),)

        datetime_format = datetime.now().strftime("%B%-d_%Y_%I%M%p")
        report_name = f"Benefit_Summary_{datetime_format}"
        start_time = start.strftime("%B %-d, %Y")
        end_time = end.strftime("%B %-d, %Y")
        title = f"(Benefits - Report({start_time} - {end_time})"
        response = download_xlsx_report(xlsx_data, report_name, title)
        return response

    base_portal = detect_base_portal_template(request)
    ctx = {
        "form": form,
        "base_reports": reverse("reports"),
        "base_template": base_portal,
    }
    return render(request, "reports/benefit_summary_report.html", ctx)