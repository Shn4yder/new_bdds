from typing import Dict, List, Set, Tuple, Sequence, DefaultDict
from datetime import datetime, timedelta, date, timezone
from collections import defaultdict

from report_generator.utils import normalize_name


def extract_resources(mpp_path: str, mapping: Dict[str, str]) -> List[str]:
    try:
        from aspose.tasks import Project  # type: ignore
    except Exception as e:
        raise RuntimeError("Не удалось импортировать aspose.tasks. Установите пакет: pip install aspose-tasks") from e
    try:
        project = Project(mpp_path)
    except Exception as e:
        raise RuntimeError(f"Не удалось открыть файл проекта: {mpp_path}") from e
    resources = []
    try:
        res_collection = getattr(project, "resources", None)
        if res_collection is None:
            return resources
        for res in res_collection:
            val = getattr(res, "name", None)
            resources.append(val if val is not None else "")
    except Exception as e:
        raise RuntimeError("Ошибка при чтении ресурсов из проекта") from e
    cleaned: List[str] = []
    seen: Set[str] = set()
    for r in resources:
        n = normalize_name(str(r))
        if not n:
            continue
        key = n.lower()
        if key not in seen:
            seen.add(key)
            cleaned.append(n)
    cleaned.sort()
    return cleaned


def _duration_to_hours(val) -> float:
    try:
        td = getattr(val, "to_timedelta", None)
        if callable(td):
            delta = td()
            return delta.total_seconds() / 3600.0
    except Exception:
        pass
    try:
        if isinstance(val, (int, float)):
            return float(val)
    except Exception:
        pass
    try:
        s = str(val)
        if s.startswith("PT"):
            hours = 0.0
            num = ""
            i = 2
            while i < len(s):
                if s[i].isdigit() or s[i] == ".":
                    num += s[i]
                else:
                    if s[i] in ("H", "h") and num:
                        hours += float(num)
                    elif s[i] in ("M", "m") and num:
                        hours += float(num) / 60.0
                    elif s[i] in ("S", "s") and num:
                        hours += float(num) / 3600.0
                    num = ""
                i += 1
            return hours
    except Exception:
        pass
    return 0.0


def _month_start(d: date) -> date:
    return date(d.year, d.month, 1)


def _add_month(d: date) -> date:
    y = d.year + (d.month // 12)
    m = (d.month % 12) + 1
    return date(y, m, 1)


def _normalize_dt(dt: datetime) -> datetime:
    if not isinstance(dt, datetime):
        return dt
    if dt.tzinfo is None:
        return dt
    return dt.astimezone(timezone.utc).replace(tzinfo=None)


def enumerate_months(start: datetime, finish: datetime) -> List[date]:
    if finish < start:
        start, finish = finish, start
    cur = _month_start(start.date())
    end = _month_start(finish.date())
    months: List[date] = []
    while cur <= end:
        months.append(cur)
        cur = _add_month(cur)
    return months


def aggregate_baseline_by_month(mpp_path: str) -> Tuple[str, List[date], List[Tuple[str, float, Dict[str, float], Dict[str, float]]]]:
    try:
        from aspose.tasks import Project  # type: ignore
    except Exception as e:
        raise RuntimeError("Не удалось импортировать aspose.tasks. Установите пакет: pip install aspose-tasks") from e
    try:
        project = Project(mpp_path)
    except Exception as e:
        raise RuntimeError(f"Не удалось открыть файл проекта: {mpp_path}") from e
    pname = getattr(project, "name", None) or getattr(getattr(project, "root_task", None), "name", None)
    if not pname:
        pname = mpp_path
    # Соберём диапазон дат задач (то, что видно в плане работ)
    t_starts: List[datetime] = []
    t_finishes: List[datetime] = []
    try:
        root = getattr(project, "root_task", None)
        stack = [root] if root is not None else []
        seen = set()
        while stack:
            t = stack.pop()
            if t is None or id(t) in seen:
                continue
            seen.add(id(t))
            s = getattr(t, "start", None)
            f = getattr(t, "finish", None)
            is_summary = False
            try:
                is_summary = bool(getattr(t, "is_summary", False))
            except Exception:
                pass
            if isinstance(s, datetime):
                s = _normalize_dt(s)
            if isinstance(f, datetime):
                f = _normalize_dt(f)
            # включаем только несводные задачи, чтобы не раздувать период
            if isinstance(s, datetime) and isinstance(f, datetime) and not is_summary:
                t_starts.append(s)
                t_finishes.append(f)
            # обход детей
            try:
                for ch in getattr(t, "children", []) or []:
                    stack.append(ch)
            except Exception:
                pass
    except Exception:
        pass
    assignments = []
    try:
        assignments = list(getattr(project, "resource_assignments", []))
    except Exception:
        assignments = []
    starts: List[datetime] = []
    finishes: List[datetime] = []
    res_map: Dict[object, Tuple[str, float]] = {}
    try:
        for res in getattr(project, "resources", []):
            name = normalize_name(getattr(res, "name", "") or "")
            rate = float(getattr(res, "standard_rate", 0.0) or 0.0)
            res_map[res] = (name, rate)
    except Exception:
        pass
    if not assignments:
        try:
            # fallback: collect assignments from tasks
            for t in getattr(getattr(project, "root_task", None), "children", []):
                for a in getattr(t, "assignments", []):
                    assignments.append(a)
        except Exception:
            assignments = []
    for a in assignments:
        s = getattr(a, "start", None)
        f = getattr(a, "finish", None)
        if isinstance(s, datetime):
            s = _normalize_dt(s)
        if isinstance(f, datetime):
            f = _normalize_dt(f)
        if isinstance(s, datetime) and isinstance(f, datetime):
            starts.append(s)
            finishes.append(f)
    # Определяем диапазон месяцев: приоритет дат задач, потом назначения, затем проектные даты
    p_start = getattr(project, "start_date", None)
    p_finish = getattr(project, "finish_date", None)
    if isinstance(p_start, datetime):
        p_start = _normalize_dt(p_start)
    if isinstance(p_finish, datetime):
        p_finish = _normalize_dt(p_finish)
    a_start = min(starts) if starts else None
    a_finish = max(finishes) if finishes else None
    t_start = min(t_starts) if t_starts else None
    t_finish = max(t_finishes) if t_finishes else None
    months: List[date] = []
    # 1) По задачам
    if isinstance(t_start, datetime) and isinstance(t_finish, datetime):
        use_start, use_finish = t_start, t_finish
        # Дополнительно ограничим проектными датами, если они уже есть и строже
        if isinstance(p_start, datetime) and isinstance(p_finish, datetime):
            inter_start = max(use_start, p_start)
            inter_finish = min(use_finish, p_finish)
            if inter_start <= inter_finish:
                use_start, use_finish = inter_start, inter_finish
        months = enumerate_months(use_start, use_finish)
    # 2) По назначениям
    elif isinstance(a_start, datetime) and isinstance(a_finish, datetime):
        use_start, use_finish = a_start, a_finish
        if isinstance(p_start, datetime) and isinstance(p_finish, datetime):
            inter_start = max(use_start, p_start)
            inter_finish = min(use_finish, p_finish)
            if inter_start <= inter_finish:
                use_start, use_finish = inter_start, inter_finish
        months = enumerate_months(use_start, use_finish)
    # 3) По проектным датам
    elif isinstance(p_start, datetime) and isinstance(p_finish, datetime):
        months = enumerate_months(p_start, p_finish)
    if not months:
        months = [date.today().replace(day=1)]
    key_months = [f"{m.year:04d}-{m.month:02d}" for m in months]
    per_resource_hours: Dict[str, DefaultDict[str, float]] = {}
    rates: Dict[str, float] = {}
    for res_obj, (rname, rate) in res_map.items():
        if rname and rname not in per_resource_hours:
            per_resource_hours[rname] = defaultdict(float)
            rates[rname] = rate
    # Построим индекс назначений по UID задачи, чтобы не зависеть от ссылочной идентичности объектов
    assigns_by_task_uid: DefaultDict[str, List[object]] = defaultdict(list)
    for a in assignments:
        try:
            t = getattr(a, "task", None)
            tuid = None
            if t is not None:
                tuid = getattr(t, "uid", None) or getattr(t, "Uid", None)
            if tuid is not None:
                assigns_by_task_uid[str(tuid)].append(a)
        except Exception:
            continue
    def iter_tasks(t):
        if t is None:
            return
        yield t
        try:
            for ch in getattr(t, "children", []) or []:
                yield from iter_tasks(ch)
        except Exception:
            return
    root = getattr(project, "root_task", None)
    for t in iter_tasks(root):
        try:
            is_summary = bool(getattr(t, "is_summary", False))
        except Exception:
            is_summary = False
        if is_summary:
            continue
        s = getattr(t, "start", None)
        f = getattr(t, "finish", None)
        if isinstance(s, datetime):
            s = _normalize_dt(s)
        if isinstance(f, datetime):
            f = _normalize_dt(f)
        total_h = _duration_to_hours(getattr(t, "work", None))
        # Соберём назначения для задачи: сначала через свойство task.assignments, затем fallback по UID
        assigns_for_task: List[object] = []
        try:
            assigns_for_task = list(getattr(t, "assignments", [])) or []
        except Exception:
            assigns_for_task = []
        if not assigns_for_task:
            try:
                tuid = getattr(t, "uid", None) or getattr(t, "Uid", None)
                if tuid is not None:
                    assigns_for_task = assigns_by_task_uid.get(str(tuid), [])
            except Exception:
                assigns_for_task = []
        if not assigns_for_task or total_h <= 0:
            continue
        # Если нет дат у задачи — возьмём минимальную/максимальную из назначений
        if not (isinstance(s, datetime) and isinstance(f, datetime)):
            s_list: List[datetime] = []
            f_list: List[datetime] = []
            for a in assigns_for_task:
                a_s = getattr(a, "start", None)
                a_f = getattr(a, "finish", None)
                if isinstance(a_s, datetime):
                    s_list.append(_normalize_dt(a_s))
                if isinstance(a_f, datetime):
                    f_list.append(_normalize_dt(a_f))
            if s_list and f_list:
                s = min(s_list)
                f = max(f_list)
        if not (isinstance(s, datetime) and isinstance(f, datetime)):
            continue
        # Вычислим веса распределения: предпочитаем assignment.work, иначе units, иначе поровну
        weights: List[float] = []
        aw_sum = 0.0
        for a in assigns_for_task:
            aw = _duration_to_hours(getattr(a, "work", None))
            weights.append(max(aw, 0.0))
            aw_sum += max(aw, 0.0)
        if aw_sum <= 0.0:
            weights = []
            w_sum = 0.0
            for a in assigns_for_task:
                u = getattr(a, "units", None)
                try:
                    u = float(u) if u is not None else 1.0
                except Exception:
                    u = 1.0
                u = max(u, 0.0)
                weights.append(u)
                w_sum += u
            aw_sum = w_sum if w_sum > 0 else float(len(assigns_for_task))
            if aw_sum <= 0.0:
                continue
        days = max((f.date() - s.date()).days + 1, 1)
        for a, w in zip(assigns_for_task, weights):
            r = getattr(a, "resource", None)
            tup = res_map.get(r, None)
            if not tup:
                rname = normalize_name(getattr(getattr(a, "resource", None), "name", "") or "")
                rate = 0.0
            else:
                rname, rate = tup
            if not rname:
                continue
            if rname not in per_resource_hours:
                per_resource_hours[rname] = defaultdict(float)
                rates[rname] = rate
            # Доля часов для назначения
            share_total = total_h * (w / aw_sum) if aw_sum > 0 else 0.0
            if share_total <= 0:
                continue
            per_day = share_total / days
            cur = s.date()
            for _ in range(days):
                k = f"{cur.year:04d}-{cur.month:02d}"
                per_resource_hours[rname][k] += per_day
                cur = cur + timedelta(days=1)
    # Fallback: если распределено 0 часов, используем назначения напрямую (timephased/assignment.work)
    try:
        total_allocated = 0.0
        for hm in per_resource_hours.values():
            for v in hm.values():
                total_allocated += float(v or 0.0)
        if total_allocated <= 0.0:
            for a in assignments:
                r = getattr(a, "resource", None)
                tup = res_map.get(r, None)
                if not tup:
                    rname = normalize_name(getattr(getattr(a, "resource", None), "name", "") or "")
                    rate = 0.0
                else:
                    rname, rate = tup
                if not rname:
                    continue
                if rname not in per_resource_hours:
                    per_resource_hours[rname] = defaultdict(float)
                    rates[rname] = rate
                # Попробуем timephased
                items = []
                try:
                    items = list(getattr(a, "timephased_data", []))
                except Exception:
                    items = []
                if items:
                    for it in items:
                        s = getattr(it, "start", None)
                        f = getattr(it, "finish", None)
                        if isinstance(s, datetime):
                            s = _normalize_dt(s)
                        if isinstance(f, datetime):
                            f = _normalize_dt(f)
                        val = getattr(it, "value", 0)
                        h = _duration_to_hours(val)
                        if not isinstance(s, datetime) or not isinstance(f, datetime):
                            continue
                        k = f"{s.year:04d}-{s.month:02d}"
                        per_resource_hours[rname][k] += h
                else:
                    s = getattr(a, "start", None)
                    f = getattr(a, "finish", None)
                    if isinstance(s, datetime):
                        s = _normalize_dt(s)
                    if isinstance(f, datetime):
                        f = _normalize_dt(f)
                    work = getattr(a, "work", None)
                    total_h = _duration_to_hours(work)
                    if isinstance(s, datetime) and isinstance(f, datetime) and total_h > 0:
                        days = max((f.date() - s.date()).days + 1, 1)
                        per_day = total_h / days
                        cur = s.date()
                        for _ in range(days):
                            k = f"{cur.year:04d}-{cur.month:02d}"
                            per_resource_hours[rname][k] += per_day
                            cur = cur + timedelta(days=1)
    except Exception:
        pass
    rows: List[Tuple[str, float, Dict[str, float], Dict[str, float]]] = []
    for rname, hours_map in per_resource_hours.items():
        rate = rates.get(rname, 0.0)
        cost_map = {k: hours_map.get(k, 0.0) * rate for k in hours_map.keys()}
        rows.append((rname, rate, dict(hours_map), cost_map))
    rows.sort(key=lambda x: x[0].lower())
    return pname, months, rows
