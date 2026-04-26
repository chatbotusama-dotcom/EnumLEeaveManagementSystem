import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import styles from './MyDashboard.module.scss';
import { sp } from '@pnp/sp/presets/all';

// ==================== INTERFACES ====================
interface ILeaveRequest {
  Id: number;
  Title: string;
  LeaveType: string;
  LeaveTypeId: number;
  LeaveDurationType: string;
  StartDate: string;
  EndDate: string;
  TotalDays: number;
  Status: string;
  EmployeeId: number;
  Country?: string;
  Reason?: string;
  HalfDayType?: string;
  Created?: string;
}

interface ILeaveBalance {
  LeaveTypeId: number;
  LeaveTypeName: string;
  TotalQuota: number;
  Utilized: number;
  Remaining: number;
}

interface IMonthlyData {
  month: string;
  monthNumber: number;
  year: number;
  days: number;
  requests: number;
}

interface ILeaveCycle {
  startDate: Date;
  endDate: Date;
  cycleNumber: number;
}

interface ILeaveQuota {
  AnnualLeaves: number;
  CasualLeaves: number;
  SickLeaves: number;
  OtherLeaves: number;
}

interface IEmployeeLeaveBalance {
  Id: number;
  LeavesBalance: number;
  Used: number;
  Remaining: number;
  Year: number;
}

// ==================== HELPER FUNCTIONS ====================
const formatDisplayDate = (dateString: string): string => {
  if (!dateString) return 'N/A';
  const date = new Date(dateString);
  return date.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  });
};

const calculateCurrentLeaveCycle = (joinDate: string): ILeaveCycle => {
  const join = new Date(joinDate);
  const today = new Date();

  let cycleStart = new Date(join);
  let cycleEnd = new Date(join);
  cycleEnd.setFullYear(cycleEnd.getFullYear() + 1);
  cycleEnd.setDate(cycleEnd.getDate() - 1);

  let cycleNumber = 1;

  while (today > cycleEnd) {
    cycleStart = new Date(cycleEnd);
    cycleStart.setDate(cycleStart.getDate() + 1);
    cycleEnd = new Date(cycleStart);
    cycleEnd.setFullYear(cycleEnd.getFullYear() + 1);
    cycleEnd.setDate(cycleEnd.getDate() - 1);
    cycleNumber++;
  }

  return { startDate: cycleStart, endDate: cycleEnd, cycleNumber };
};

const getLeaveQuota = async (countryId: number, countryCodeId: number): Promise<ILeaveQuota> => {
  try {
    const items = await sp.web.lists
      .getByTitle('Leave Quota')
      .items
      .select('AnnualLeaves', 'CasualLeaves', 'SickLeaves', 'OtherLeaves')
      .filter(`Country/Id eq ${countryId} and CountryCode/Id eq ${countryCodeId}`)
      .top(1)
      .get();

    if (items && items.length > 0) {
      return {
        AnnualLeaves: items[0].AnnualLeaves || 0,
        CasualLeaves: items[0].CasualLeaves || 0,
        SickLeaves: items[0].SickLeaves || 0,
        OtherLeaves: items[0].OtherLeaves || 0
      };
    }
  } catch (error) {
    console.error('[Quota] Error:', error);
  }
  
  return { AnnualLeaves: 0, CasualLeaves: 0, SickLeaves: 0, OtherLeaves: 0 };
};

// ==================== MAIN COMPONENT ====================
const MyDashboard: React.FC = () => {
  const [leaveRequests, setLeaveRequests] = useState<ILeaveRequest[]>([]);
  const [filteredRequests, setFilteredRequests] = useState<ILeaveRequest[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [leaveBalances, setLeaveBalances] = useState<ILeaveBalance[]>([]);
  const [selectedRequest, setSelectedRequest] = useState<ILeaveRequest | null>(null);
  const [showModal, setShowModal] = useState<boolean>(false);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [itemsPerPage] = useState<number>(5);
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [yearFilter, setYearFilter] = useState<string>('all');
  const [monthlyData, setMonthlyData] = useState<IMonthlyData[]>([]);
  const [selectedYear, setSelectedYear] = useState<number>(new Date().getFullYear());
  const [availableYears, setAvailableYears] = useState<number[]>([]);
  const [, setEmployeeLeaveBalance] = useState<IEmployeeLeaveBalance | null>(null);

  useEffect(() => {
    loadInitialData();
  }, []);

  useEffect(() => {
    filterRequests();
    setCurrentPage(1);
  }, [statusFilter, yearFilter, leaveRequests]);

  useEffect(() => {
    if (leaveRequests.length > 0 && selectedYear) {
      calculateMonthlyTrends(leaveRequests, selectedYear);
    }
  }, [leaveRequests, selectedYear]);

  // ==================== MAIN DATA LOADING ====================
  const loadInitialData = async () => {
    try {
      setLoading(true);
      const user = await sp.web.currentUser();
      const currentYear = new Date().getFullYear();

      // Parallelize independent list fetches to reduce wait time
      const [empData, types, requests, countries, empBalance] = await Promise.all([
        sp.web.lists
          .getByTitle("Employee Information List")
          .items
          .select("Id", "JoiningDate", "EmploymentType", "Country/Id", "CountryCode/Id", "EmployeeName/Id", "EmployeeName/Title", "Department/Title")
          .expand("Country", "CountryCode", "EmployeeName", "Department")
          .filter(`EmployeeName/EMail eq '${user.Email}'`)
          .get(),
        sp.web.lists.getByTitle("Leave Types").items.select("Id", "Title").get(),
        sp.web.lists
          .getByTitle("Leave Requests")
          .items
          .select("Id", "Title", "LeaveTypeId", "LeaveDurationType", "StartDate", "EndDate", "TotalDays", "Status", "HalfDayType", "CountryId", "Reason", "EmployeeId", "Created")
          .filter(`EmployeeId eq ${user.Id}`)
          .orderBy("StartDate", false)
          .get(),
        sp.web.lists.getByTitle("Country").items.select("Id", "Title").get(),
        sp.web.lists
          .getByTitle('Employee Leave Balance')
          .items
          .select('Id', 'LeavesBalance', 'Used', 'Remaining', 'Year', 'Employee/Id', 'Employee/Title', 'Employee/EMail')
          .expand('Employee')
          .filter(`Employee/EMail eq '${user.Email}' and Year eq ${currentYear}`)
          .top(1)
          .get()
      ]);

      let joinDate = '';
      let countryId = 0;
      let countryCodeId = 0;

      if (empData && empData.length > 0) {
        joinDate = empData[0].JoiningDate;
        countryId = empData[0].Country?.Id || 0;
        countryCodeId = empData[0].CountryCode?.Id || 0;
      }

      // Format requests
      const formattedRequests: ILeaveRequest[] = (requests || []).map(request => {
        const leaveType = (types || []).find(lt => lt.Id === request.LeaveTypeId);
        const country = (countries || []).find(c => c.Id === request.CountryId);
        return {
          Id: request.Id,
          Title: request.Title,
          LeaveType: leaveType?.Title || 'Unknown',
          LeaveTypeId: request.LeaveTypeId,
          LeaveDurationType: request.LeaveDurationType,
          StartDate: request.StartDate,
          EndDate: request.EndDate,
          TotalDays: request.TotalDays,
          Status: request.Status,
          EmployeeId: request.EmployeeId,
          Country: country?.Title || 'Unknown',
          Reason: request.Reason,
          HalfDayType: request.HalfDayType,
          Created: request.Created
        };
      });

      setLeaveRequests(formattedRequests);

      // Process employee balance if present
      let balance: IEmployeeLeaveBalance | null = null;
      if (empBalance && empBalance.length > 0) {
        balance = {
          Id: empBalance[0].Id,
          LeavesBalance: empBalance[0].LeavesBalance,
          Used: empBalance[0].Used,
          Remaining: empBalance[0].Remaining,
          Year: empBalance[0].Year
        };
        setEmployeeLeaveBalance(balance);
      }

      // Calculate current cycle and balances (only if we have joinDate)
      if (joinDate) {
        const cycle = calculateCurrentLeaveCycle(joinDate);
        const quota = await getLeaveQuota(countryId, countryCodeId);
        await calculateLeaveBalances(cycle, quota, formattedRequests, balance);
      }

      // Get available years for filter
      const joinYear = joinDate ? new Date(joinDate).getFullYear() : new Date().getFullYear();
      const years = [];
      for (let i = joinYear; i <= currentYear; i++) {
        years.push(i);
      }
      setAvailableYears(years.sort((a, b) => b - a));

    } catch (error) {
      console.error("[Dashboard] Error loading data:", error);
    } finally {
      setLoading(false);
    }
  };

  // ==================== LEAVE BALANCE CALCULATION ====================
  const calculateLeaveBalances = async (
    cycle: ILeaveCycle, 
    quota: ILeaveQuota, 
    requests: ILeaveRequest[],
    employeeLeaveBalance: IEmployeeLeaveBalance | null
  ) => {
    const cycleNumber = cycle.cycleNumber;
    
    // Filter approved leaves within current cycle
    const cycleRequests = requests.filter(req => {
      const startDate = new Date(req.StartDate);
      const isApproved = req.Status === 'Approved';
      const inCycle = startDate >= cycle.startDate && startDate <= cycle.endDate;
      return isApproved && inCycle;
    });
    
    // Calculate used leaves by type
    let usedAnnual = 0, usedCasual = 0, usedSick = 0, usedOther = 0;
    
    cycleRequests.forEach(req => {
      const days = req.TotalDays || 0;
      switch (req.LeaveTypeId) {
        case 1: usedAnnual += days; break;
        case 2: usedCasual += days; break;
        case 3: usedSick += days; break;
        case 4: usedOther += days; break;
      }
    });
    
    // ========== ANNUAL LEAVE CALCULATION (No Pro-rata) ==========
    let totalAnnual = quota.AnnualLeaves;
    
    // Apply carry forward from Employee Leave Balance for cycle #2+
    if (cycleNumber > 1 && employeeLeaveBalance) {
      totalAnnual = quota.AnnualLeaves + (employeeLeaveBalance.Remaining || 0);
    }
    
    totalAnnual = Math.max(0, totalAnnual - usedAnnual);
    
    // ========== CASUAL LEAVE CALCULATION (No Pro-rata) ==========
    let totalCasual = quota.CasualLeaves;
    totalCasual = Math.max(0, totalCasual - usedCasual);
    
    // ========== SICK LEAVE CALCULATION (No Pro-rata) ==========
    let totalSick = quota.SickLeaves;
    totalSick = Math.max(0, totalSick - usedSick);
    
    // ========== OTHER LEAVE CALCULATION ==========
    let totalOther = quota.OtherLeaves;
    totalOther = Math.max(0, totalOther - usedOther);
    
    // Build balances array
    const balances: ILeaveBalance[] = [
      {
        LeaveTypeId: 1,
        LeaveTypeName: 'Annual Leave',
        TotalQuota: totalAnnual + usedAnnual,
        Utilized: usedAnnual,
        Remaining: totalAnnual
      },
      {
        LeaveTypeId: 2,
        LeaveTypeName: 'Casual Leave',
        TotalQuota: totalCasual + usedCasual,
        Utilized: usedCasual,
        Remaining: totalCasual
      },
      {
        LeaveTypeId: 3,
        LeaveTypeName: 'Sick Leave',
        TotalQuota: totalSick + usedSick,
        Utilized: usedSick,
        Remaining: totalSick
      },
      {
        LeaveTypeId: 4,
        LeaveTypeName: 'Other Leave',
        TotalQuota: totalOther + usedOther,
        Utilized: usedOther,
        Remaining: totalOther
      }
    ];
    
    setLeaveBalances(balances);
  };

  // ==================== MONTHLY TRENDS ====================
  const calculateMonthlyTrends = (requests: ILeaveRequest[], year: number) => {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthlyDays = new Array(12).fill(0);
    const monthlyRequestCount = new Array(12).fill(0);
    
    const yearRequests = requests.filter(r => {
      const startDate = new Date(r.StartDate);
      return r.Status === 'Approved' && startDate.getFullYear() === year;
    });
    
    yearRequests.forEach(request => {
      const startDate = new Date(request.StartDate);
      const endDate = new Date(request.EndDate);
      const startMonth = startDate.getMonth();
      const endMonth = endDate.getMonth();
      
      if (startMonth === endMonth) {
        monthlyDays[startMonth] += request.TotalDays;
        monthlyRequestCount[startMonth] += 1;
      } else {
        const daysInStartMonth = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 0).getDate();
        const daysInStart = daysInStartMonth - startDate.getDate() + 1;
        const daysInEnd = endDate.getDate();
        
        monthlyDays[startMonth] += daysInStart;
        monthlyRequestCount[startMonth] += 1;
        monthlyDays[endMonth] += daysInEnd;
        monthlyRequestCount[endMonth] += 1;
      }
    });
    
    const monthlyData: IMonthlyData[] = months.map((month, index) => ({
      month,
      monthNumber: index,
      year,
      days: monthlyDays[index],
      requests: monthlyRequestCount[index]
    }));
    
    setMonthlyData(monthlyData);
  };

  // ==================== FILTER FUNCTIONS ====================
  const filterRequests = () => {
    let filtered = [...leaveRequests];
    if (statusFilter !== 'all') {
      filtered = filtered.filter(r => r.Status === statusFilter);
    }
    if (yearFilter !== 'all') {
      filtered = filtered.filter(r => {
        const startDate = new Date(r.StartDate);
        return startDate.getFullYear().toString() === yearFilter;
      });
    }
    setFilteredRequests(filtered);
  };

  const handleYearChange = (year: number) => {
    setSelectedYear(year);
  };

  const getAvailableYearsForFilter = useMemo(() => {
    const years = leaveRequests.map(r => new Date(r.StartDate).getFullYear());
    return Array.from(new Set(years)).sort((a, b) => b - a);
  }, [leaveRequests]);

  // ==================== UI HELPERS ====================
  const formatDateTime = (dateString: string) => {
    if (!dateString) return 'N/A';
    const date = new Date(dateString);
    return date.toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  };

  const getStatusClass = (status: string) => {
    switch(status?.toLowerCase()) {
      case 'pending':
      case 'pending on manager':
      case 'pending on manager':
      case 'pending on hr':
      case 'pending on executive':
        return styles.statusPending;
      case 'approved':
        return styles.statusApproved;
      case 'rejected':
      case 'rejected by manager':
      case 'rejected by hr':
      case 'rejected by executive':
      case 'send back by manager':
      case 'send back by hr':
      case 'send back by executive':
        return styles.statusRejected;
      default:
        return '';
    }
  };

  const getDurationDisplay = (request: ILeaveRequest) => {
    if (request.LeaveDurationType === 'Half Day') {
      return `${request.LeaveDurationType} (${request.HalfDayType || 'N/A'})`;
    }
    return request.LeaveDurationType;
  };

  const handleViewRequest = (request: ILeaveRequest) => {
    setSelectedRequest(request);
    setShowModal(true);
  };

  const closeModal = () => {
    setShowModal(false);
    setSelectedRequest(null);
  };

  // ==================== PAGINATION ====================
  const indexOfLastItem = currentPage * itemsPerPage;
  const indexOfFirstItem = indexOfLastItem - itemsPerPage;
  const currentItems = useMemo(() => filteredRequests.slice(indexOfFirstItem, indexOfLastItem), [filteredRequests, indexOfFirstItem, indexOfLastItem]);
  const totalPages = useMemo(() => Math.ceil(filteredRequests.length / itemsPerPage), [filteredRequests.length, itemsPerPage]);
  const paginate = (pageNumber: number) => setCurrentPage(pageNumber);

  // ==================== BALANCE GETTERS ====================
  const getBalance = (name: string) => leaveBalances.find(b => b.LeaveTypeName === name);
  
  const annualBalance = useMemo(() => getBalance('Annual Leave'), [leaveBalances]);
  const casualBalance = useMemo(() => getBalance('Casual Leave'), [leaveBalances]);
  const sickBalance = useMemo(() => getBalance('Sick Leave'), [leaveBalances]);
  const otherBalance = useMemo(() => getBalance('Other Leave'), [leaveBalances]);

  const maxDays = useMemo(() => Math.max(...monthlyData.map(m => m.days), 5), [monthlyData]);

  // Donut chart: Utilized (Red) + Remaining (Green) = Total
  const getDonutSegments = (total: number, utilized: number) => {
    const remaining = Math.max(0, total - utilized);
    const utilizedPercent = total > 0 ? (utilized / total) * 283 : 0;
    const remainingPercent = total > 0 ? (remaining / total) * 283 : 0;
    return { utilizedPercent, remainingPercent };
  };

  return (
    <div className={styles.dashboard}>
      {loading && leaveRequests.length === 0 && (
        <div className={styles.fullPageLoader}>
          <div className={styles.fullSpinner}></div>
        </div>
      )}
      {/* 4 Donut Charts */}
      <div className={styles.donutGrid}>
        {/* Annual Leave Card */}
        <div className={styles.donutCard}>
          <h3>Annual Leave</h3>
          <div className={styles.donutWrapper}>
            <div className={styles.donutSmall}>
              <svg viewBox="0 0 100 100" className={styles.donutSvgSmall}>
                <circle cx="50" cy="50" r="45" fill="none" stroke="#e2e8f0" strokeWidth="10"/>
                <circle 
                  cx="50" cy="50" r="45" fill="none" 
                  stroke="#ef4444" 
                  strokeWidth="10"
                  strokeDasharray={`${getDonutSegments(annualBalance?.TotalQuota || 0, annualBalance?.Utilized || 0).utilizedPercent} 283`}
                  transform="rotate(-90 50 50)"
                />
                {((annualBalance?.Remaining || 0) > 0) && (
                  <circle 
                    cx="50" cy="50" r="45" fill="none" 
                    stroke="#10b981" 
                    strokeWidth="10"
                    strokeDasharray={`${getDonutSegments(annualBalance?.TotalQuota || 0, annualBalance?.Utilized || 0).remainingPercent} 283`}
                    strokeDashoffset={`-${getDonutSegments(annualBalance?.TotalQuota || 0, annualBalance?.Utilized || 0).utilizedPercent}`}
                    transform="rotate(-90 50 50)"
                  />
                )}
              </svg>
              <div className={styles.donutCenterSmall}>
                <div className={styles.donutTotalSmall}>{annualBalance?.TotalQuota || 0}</div>
                <div className={styles.donutLabelSmall}>Total</div>
              </div>
            </div>
            <div className={styles.donutStats}>
              <div className={styles.statRow}>
                <span className={styles.utilizedDot}></span>
                <span>Utilized:</span>
                <strong className={styles.utilizedText}>{annualBalance?.Utilized || 0} days</strong>
              </div>
              <div className={styles.statRow}>
                <span className={styles.remainingDot}></span>
                <span>Remaining:</span>
                <strong className={styles.remainingText}>{annualBalance?.Remaining || 0} days</strong>
              </div>
            </div>
          </div>
        </div>

        {/* Casual Leave Card */}
        <div className={styles.donutCard}>
          <h3>Casual Leave</h3>
          <div className={styles.donutWrapper}>
            <div className={styles.donutSmall}>
              <svg viewBox="0 0 100 100" className={styles.donutSvgSmall}>
                <circle cx="50" cy="50" r="45" fill="none" stroke="#e2e8f0" strokeWidth="10"/>
                <circle 
                  cx="50" cy="50" r="45" fill="none" 
                  stroke="#ef4444" 
                  strokeWidth="10"
                  strokeDasharray={`${getDonutSegments(casualBalance?.TotalQuota || 0, casualBalance?.Utilized || 0).utilizedPercent} 283`}
                  transform="rotate(-90 50 50)"
                />
                {((casualBalance?.Remaining || 0) > 0) && (
                  <circle 
                    cx="50" cy="50" r="45" fill="none" 
                    stroke="#10b981" 
                    strokeWidth="10"
                    strokeDasharray={`${getDonutSegments(casualBalance?.TotalQuota || 0, casualBalance?.Utilized || 0).remainingPercent} 283`}
                    strokeDashoffset={`-${getDonutSegments(casualBalance?.TotalQuota || 0, casualBalance?.Utilized || 0).utilizedPercent}`}
                    transform="rotate(-90 50 50)"
                  />
                )}
              </svg>
              <div className={styles.donutCenterSmall}>
                <div className={styles.donutTotalSmall}>{casualBalance?.TotalQuota || 0}</div>
                <div className={styles.donutLabelSmall}>Total</div>
              </div>
            </div>
            <div className={styles.donutStats}>
              <div className={styles.statRow}>
                <span className={styles.utilizedDot}></span>
                <span>Utilized:</span>
                <strong className={styles.utilizedText}>{casualBalance?.Utilized || 0} days</strong>
              </div>
              <div className={styles.statRow}>
                <span className={styles.remainingDot}></span>
                <span>Remaining:</span>
                <strong className={styles.remainingText}>{casualBalance?.Remaining || 0} days</strong>
              </div>
            </div>
          </div>
        </div>

        {/* Sick Leave Card */}
        <div className={styles.donutCard}>
          <h3>Sick Leave</h3>
          <div className={styles.donutWrapper}>
            <div className={styles.donutSmall}>
              <svg viewBox="0 0 100 100" className={styles.donutSvgSmall}>
                <circle cx="50" cy="50" r="45" fill="none" stroke="#e2e8f0" strokeWidth="10"/>
                <circle 
                  cx="50" cy="50" r="45" fill="none" 
                  stroke="#ef4444" 
                  strokeWidth="10"
                  strokeDasharray={`${getDonutSegments(sickBalance?.TotalQuota || 0, sickBalance?.Utilized || 0).utilizedPercent} 283`}
                  transform="rotate(-90 50 50)"
                />
                {((sickBalance?.Remaining || 0) > 0) && (
                  <circle 
                    cx="50" cy="50" r="45" fill="none" 
                    stroke="#10b981" 
                    strokeWidth="10"
                    strokeDasharray={`${getDonutSegments(sickBalance?.TotalQuota || 0, sickBalance?.Utilized || 0).remainingPercent} 283`}
                    strokeDashoffset={`-${getDonutSegments(sickBalance?.TotalQuota || 0, sickBalance?.Utilized || 0).utilizedPercent}`}
                    transform="rotate(-90 50 50)"
                  />
                )}
              </svg>
              <div className={styles.donutCenterSmall}>
                <div className={styles.donutTotalSmall}>{sickBalance?.TotalQuota || 0}</div>
                <div className={styles.donutLabelSmall}>Total</div>
              </div>
            </div>
            <div className={styles.donutStats}>
              <div className={styles.statRow}>
                <span className={styles.utilizedDot}></span>
                <span>Utilized:</span>
                <strong className={styles.utilizedText}>{sickBalance?.Utilized || 0} days</strong>
              </div>
              <div className={styles.statRow}>
                <span className={styles.remainingDot}></span>
                <span>Remaining:</span>
                <strong className={styles.remainingText}>{sickBalance?.Remaining || 0} days</strong>
              </div>
            </div>
          </div>
        </div>

        {/* Other Leave Card */}
        <div className={styles.donutCard}>
          <h3>Other Leave</h3>
          <div className={styles.donutWrapper}>
            <div className={styles.donutSmall}>
              <svg viewBox="0 0 100 100" className={styles.donutSvgSmall}>
                <circle cx="50" cy="50" r="45" fill="none" stroke="#e2e8f0" strokeWidth="10"/>
                <circle 
                  cx="50" cy="50" r="45" fill="none" 
                  stroke="#ef4444" 
                  strokeWidth="10"
                  strokeDasharray={`${getDonutSegments(otherBalance?.TotalQuota || 0, otherBalance?.Utilized || 0).utilizedPercent} 283`}
                  transform="rotate(-90 50 50)"
                />
                {((otherBalance?.Remaining || 0) > 0) && (
                  <circle 
                    cx="50" cy="50" r="45" fill="none" 
                    stroke="#10b981" 
                    strokeWidth="10"
                    strokeDasharray={`${getDonutSegments(otherBalance?.TotalQuota || 0, otherBalance?.Utilized || 0).remainingPercent} 283`}
                    strokeDashoffset={`-${getDonutSegments(otherBalance?.TotalQuota || 0, otherBalance?.Utilized || 0).utilizedPercent}`}
                    transform="rotate(-90 50 50)"
                  />
                )}
              </svg>
              <div className={styles.donutCenterSmall}>
                <div className={styles.donutTotalSmall}>{otherBalance?.TotalQuota || 0}</div>
                <div className={styles.donutLabelSmall}>Total</div>
              </div>
            </div>
            <div className={styles.donutStats}>
              <div className={styles.statRow}>
                <span className={styles.utilizedDot}></span>
                <span>Utilized:</span>
                <strong className={styles.utilizedText}>{otherBalance?.Utilized || 0} days</strong>
              </div>
              <div className={styles.statRow}>
                <span className={styles.remainingDot}></span>
                <span>Remaining:</span>
                <strong className={styles.remainingText}>{otherBalance?.Remaining || 0} days</strong>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Color Legend */}
      <div className={styles.legendBar}>
        <div className={styles.legendItem}>
          <div className={styles.legendColor} style={{ background: '#10b981' }}></div>
          <span>Remaining Balance</span>
        </div>
        <div className={styles.legendItem}>
          <div className={styles.legendColor} style={{ background: '#ef4444' }}></div>
          <span>Utilized Balance</span>
        </div>
      </div>

      {/* Monthly Trends Chart */}
      <div className={styles.chartCardLarge}>
        <div className={styles.chartHeader}>
          <h3>Monthly Leave Trends</h3>
          <select 
            value={selectedYear} 
            onChange={(e) => handleYearChange(Number(e.target.value))}
            className={styles.yearSelect}
          >
            {availableYears.map(year => (
              <option key={year} value={year}>{year}</option>
            ))}
          </select>
        </div>
        <div className={styles.monthlyChartContainer}>
          <div className={styles.monthlyChart}>
            {monthlyData.map((data, index) => {
              const height = maxDays > 0 ? (data.days / maxDays) * 140 : 0;
              return (
                <div key={index} className={styles.monthBarWrapper}>
                  <div className={styles.monthBarColumn}>
                    <div 
                      className={styles.monthBarFill} 
                      style={{ height: `${Math.max(height, 4)}px` }}
                      title={`${data.month}: ${data.days} days (${data.requests} requests)`}
                    >
                      {data.days > 0 && <span className={styles.monthBarValue}>{data.days}</span>}
                    </div>
                  </div>
                  <div className={styles.monthBarLabel}>{data.month}</div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* My Requests Section */}
      <div className={styles.requestsSection}>
        <div className={styles.sectionHeader}>
          <h3>My Leave Requests</h3>
          <div className={styles.simpleFilters}>
            <select 
              value={statusFilter} 
              onChange={(e) => setStatusFilter(e.target.value)}
              className={styles.filterSelect}
            >
              <option value="all">All Status</option>
              <option value="Pending on Manager">Pending on Manager</option>
              <option value="Pending on HR">Pending on HR</option>
              <option value="Pending on Executive">Pending on Executive</option>
              <option value="Rejected by Manager">Rejected by Manager</option>
              <option value="Rejected by HR">Rejected by HR</option>
              <option value="Rejected by Executive">Rejected by Executive</option>
              <option value="Send Back by Manager">Send Back by Manager</option>
              <option value="Send Back by HR">Send Back by HR</option>
              <option value="Send Back by Executive">Send Back by Executive</option>
              <option value="Approved">Approved</option>
            </select>
            <select 
              value={yearFilter} 
              onChange={(e) => setYearFilter(e.target.value)}
              className={styles.filterSelect}
            >
              <option value="all">All Years</option>
              {getAvailableYearsForFilter.map(year => (
                <option key={year} value={year}>{year}</option>
              ))}
            </select>
          </div>
        </div>

        <div className={styles.tableWrapper}>
          {loading ? (
            <div className={styles.loadingState}>
              <div className={styles.spinner}></div>
              <p>Loading your requests...</p>
            </div>
          ) : currentItems.length > 0 ? (
            <>
              <table className={styles.requestTable}>
                <thead>
                  <tr>
                    <th>Leave Type</th>
                    <th>Duration</th>
                    <th>Start Date</th>
                    <th>End Date</th>
                    <th>Days</th>
                    <th>Status</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {currentItems.map(request => (
                    <tr key={request.Id}>
                      <td className={styles.leaveTypeCell}>{request.LeaveType}</td>
                      <td>
                        <span className={styles.durationBadge}>
                          {getDurationDisplay(request)}
                        </span>
                      </td>
                      <td>{formatDisplayDate(request.StartDate)}</td>
                      <td>{formatDisplayDate(request.EndDate)}</td>
                      <td>
                        <span className={styles.daysBadge}>{request.TotalDays}</span>
                      </td>
                      <td>
                        <span className={`${styles.statusBadge} ${getStatusClass(request.Status)}`}>
                          {request.Status}
                        </span>
                      </td>
                      <td>
                        <button 
                          className={styles.viewBtn}
                          onClick={() => handleViewRequest(request)}
                        >
                          View Details
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {totalPages > 1 && (
                <div className={styles.pagination}>
                  <button
                    onClick={() => paginate(currentPage - 1)}
                    disabled={currentPage === 1}
                    className={styles.pageBtn}
                  >
                    Previous
                  </button>
                  <span className={styles.pageInfo}>
                    Page {currentPage} of {totalPages}
                  </span>
                  <button
                    onClick={() => paginate(currentPage + 1)}
                    disabled={currentPage === totalPages}
                    className={styles.pageBtn}
                  >
                    Next
                  </button>
                </div>
              )}
            </>
          ) : (
            <div className={styles.emptyState}>
              <svg width="64" height="64" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M9 12H15M9 16H12M12 8H12.01M3 12C3 13.1819 3.23279 14.3522 3.68508 15.4442C4.13738 16.5361 4.80031 17.5282 5.63604 18.364C6.47177 19.1997 7.46392 19.8626 8.55585 20.3149C9.64778 20.7672 10.8181 21 12 21C13.1819 21 14.3522 20.7672 15.4442 20.3149C16.5361 19.8626 17.5282 19.1997 18.364 18.364C19.1997 17.5282 19.8626 16.5361 20.3149 15.4442C20.7672 14.3522 21 13.1819 21 12C21 9.61305 20.0518 7.32387 18.364 5.63604C16.6761 3.94821 14.3869 3 12 3C9.61305 3 7.32387 3.94821 5.63604 5.63604C3.94821 7.32387 3 9.61305 3 12Z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
              </svg>
              <p>No leave requests found</p>
              <span>Your leave requests will appear here once you submit them</span>
            </div>
          )}
        </div>
      </div>

      {/* Modal */}
      {showModal && selectedRequest && (
        <div className={styles.modalOverlay} onClick={closeModal}>
          <div className={styles.modalContent} onClick={(e) => e.stopPropagation()}>
            <div className={styles.modalHeader}>
              <h3>Leave Request Details</h3>
              <button className={styles.modalClose} onClick={closeModal}>×</button>
            </div>
            <div className={styles.modalBody}>
              <div className={styles.detailRow}>
                <div className={styles.detailLabel}>Leave Type:</div>
                <div className={styles.detailValue}>{selectedRequest.LeaveType}</div>
              </div>
              <div className={styles.detailRow}>
                <div className={styles.detailLabel}>Duration:</div>
                <div className={styles.detailValue}>{getDurationDisplay(selectedRequest)}</div>
              </div>
              <div className={styles.detailRow}>
                <div className={styles.detailLabel}>Start Date:</div>
                <div className={styles.detailValue}>{formatDisplayDate(selectedRequest.StartDate)}</div>
              </div>
              <div className={styles.detailRow}>
                <div className={styles.detailLabel}>End Date:</div>
                <div className={styles.detailValue}>{formatDisplayDate(selectedRequest.EndDate)}</div>
              </div>
              <div className={styles.detailRow}>
                <div className={styles.detailLabel}>Total Days:</div>
                <div className={styles.detailValue}>{selectedRequest.TotalDays}</div>
              </div>
              <div className={styles.detailRow}>
                <div className={styles.detailLabel}>Status:</div>
                <div className={styles.detailValue}>
                  <span className={`${styles.statusBadge} ${getStatusClass(selectedRequest.Status)}`}>
                    {selectedRequest.Status}
                  </span>
                </div>
              </div>
              <div className={styles.detailRow}>
                <div className={styles.detailLabel}>Country:</div>
                <div className={styles.detailValue}>{selectedRequest.Country || 'N/A'}</div>
              </div>
              <div className={styles.detailRow}>
                <div className={styles.detailLabel}>Submitted On:</div>
                <div className={styles.detailValue}>{formatDateTime(selectedRequest.Created || '')}</div>
              </div>
              <div className={styles.detailRow}>
                <div className={styles.detailLabel}>Reason:</div>
                <div className={styles.detailValue}>{selectedRequest.Reason || 'N/A'}</div>
              </div>
            </div>
            <div className={styles.modalFooter}>
              <button className={styles.closeModalBtn} onClick={closeModal}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default React.memo(MyDashboard);