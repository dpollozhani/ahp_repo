import ahpy

criteria_comparisons = {
    ('direct costs', 'system maintenance'):3,
    ('direct costs', 'accessibility'):1,
    ('direct costs', 'self service'):3,
    ('direct costs', 'etl process'):1/3,
    ('direct costs', 'security'):1,
    ('system maintenance', 'accessibility'):1/3,
    ('system maintenance', 'self service'):3,
    ('system maintenance', 'etl process'):1,
    ('system maintenance', 'security'):1/3,
    ('self service', 'etl process'):1/3,
    ('self service', 'security'):1/3,
    ('etl process', 'security'):1,
    ('accessibility', 'self service'):3,
    ('accessibility', 'etl process'):1,
    ('accessibility', 'security'):1
    }

criteria = ahpy.Compare('Criteria', criteria_comparisons)
report = criteria.report(show=True)