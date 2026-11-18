import urllib.request
import json

def test():
    # Step 1
    req1 = urllib.request.Request('http://localhost:8080/api/run-skill-1', method='POST')
    res1 = json.loads(urllib.request.urlopen(req1).read())
    print('Step 1 (Cleansing):', res1['status'], '| Valid report txs:', res1['summary']['valid_report_period'], '| Unqualified:', res1['unqualified_count'])

    # HITL Approve
    body2 = json.dumps({'decision': 'APPROVE', 'notes': 'Data verified and approved by analyst.'}).encode('utf-8')
    req2 = urllib.request.Request('http://localhost:8080/api/human-verify', data=body2, headers={'Content-Type': 'application/json'}, method='POST')
    res2 = json.loads(urllib.request.urlopen(req2).read())
    print('Step 2 (HITL Gate):', res2['status'], '| Message:', res2['message'])

    # Step 3
    req3 = urllib.request.Request('http://localhost:8080/api/run-skill-2', method='POST')
    res3 = json.loads(urllib.request.urlopen(req3).read())
    print('Step 3 (Skill 2):', res3['status'])
    print('Top 3 Stores Identified:')
    for s in res3['top_3_stores']:
        print(f"  Rank #{s['rank']}: {s['store_name']} ({s['store_id']}) - Risk Score: {s['risk_score']} | Net Profit: {s['net_profit']:,.0f} VND | Target %: {s['rev_target_pct']}% | Return Rate: {s['return_rate_pct']}%")

    # Fetch Dashboard
    req4 = urllib.request.Request('http://localhost:8080/api/dashboard')
    dash = json.loads(urllib.request.urlopen(req4).read())
    print('Dashboard KPIs:', dash['kpis'])

if __name__ == '__main__':
    test()

