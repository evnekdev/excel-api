"""Summarize the #122 assessment by source kind and independent evidence axis."""
import collections
import json
import subprocess
import sys

import validate_rust_draft as audit


def observed_count(items, key):
    return sum(item[key] != 'unassessed' and not (key == 'test_state' and item[key] == 'test-source-present') for item in items)


def main():
    audit.verify_current_inputs()
    catalogue = {item['id']: item for item in map(json.loads, audit.git_show(audit.CATALOGUE).splitlines())}
    rows = audit.catalogue_validator.read_jsonl(audit.DRAFT.read_bytes())
    audit.verify(rows)
    kinds = collections.defaultdict(list)
    for row in rows:
        kinds[catalogue[row['capi_id']]['kind']].append(row)
    print('| Source kind | IDs | Raw observations / unknown | Function observations / unknown | Test passed / unknown | Host assessed / unknown |')
    print('| --- | ---: | ---: | ---: | ---: | ---: |')
    for kind, items in sorted(kinds.items()):
        count = len(items)
        values = []
        for key in ('raw_state', 'functional_state', 'test_state', 'host_state'):
            observed = observed_count(items, key)
            values.append(f'{observed} / {count - observed}')
        print(f'| {kind} | {count} | ' + ' | '.join(values) + ' |')


if __name__ == '__main__':
    try:
        main()
    except (audit.catalogue_validator.HistoryUnavailable, audit.catalogue_validator.SourceUnavailable) as exc:
        print(f'NOT VERIFIED: {exc}', file=sys.stderr)
        sys.exit(3)
    except (ValueError, OSError, subprocess.CalledProcessError) as exc:
        print(f'FAIL: assessment cannot be summarized: {exc}', file=sys.stderr)
        sys.exit(1)
