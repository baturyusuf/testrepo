"""Backward entrypoint kept for repository users; adaptive tests are authoritative."""
from test_adaptive_layout import test_all_combinations, test_workbook_and_bridge

if __name__ == '__main__':
    print(f'{test_all_combinations()} layout combinations, '
          f'{test_workbook_and_bridge()} Excel/PDF manifests passed')