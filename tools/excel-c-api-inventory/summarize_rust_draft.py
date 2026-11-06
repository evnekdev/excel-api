"""Summarize the #122 assessment by source kind and independent evidence axis."""
import collections
import json

import validate_rust_draft as audit


def main():
    audit.verify_current_inputs()
    catalogue = {item['id']: item for item in map(json.loads, audit.git_show(audit.CATALOGUE).splitlines())}
    rows = [json.loads(line) for line in audit.DRAFT.read_text(encoding='utf8').splitlines()]
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
            unknown = sum(item[key] == 'unassessed' for item in items)
            values.append(f'{count - unknown} / {unknown}')
        print(f'| {kind} | {count} | ' + ' | '.join(values) + ' |')


if __name__ == '__main__':
    main()
