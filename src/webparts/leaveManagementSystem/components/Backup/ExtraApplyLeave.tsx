import * as React from 'react';
import styles from './ExtraApplyLeave.module.scss';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import EditLeave from './ExtraEditLeave';

interface IFormState {
  leaveTypeId: number;
  leaveDurationType: 'Full Day' | 'Half Day';
  startDate: string;
  endDate: string;
  totalDays: number;
  halfDayType: 'First Half' | 'Second Half';
  reason: string;
  attachment: File | null;
  otherLeaveType: string;
}

interface IFormErrors {
  leaveType?: string;
  startDate?: string;
  endDate?: string;
  reason?: string;
  overlap?: string;
  attachment?: string;
  submit?: string;
  weekend?: string;
  manager?: string;
  probation?: string;
  otherLeaveType?: string;
  insufficientQuota?: string;
}

interface IEmployeeData {
  id: number;
  name: string;
  department: string;
  departmentId: number;
  country: string;
  countryId: number;
  countryCode: string;
  countryCodeId: number;
  managerId: number;
  managerName: string;
  managerEmail: string;
  joinDate: string;
  employmentType: string;
  userImageUrl: string;
}

interface ILeaveType {
  Id: number;
  Title: string;
}

interface ILeaveQuota {
  Id: number;
  CountryId: number;
  CountryCodeId: number;
  Leaves: number;
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
  Year: string;
}

interface ILeaveCycle {
  startDate: Date;
  endDate: Date;
  cycleNumber: number;
}

interface ILeaveRequest {
  Id: number;
  StartDate: string;
  EndDate: string;
  Status: string;
  TotalDays: number;
  LeaveTypeId: number;
}

interface IProRataResult {
  proRataAnnualLeave: number;
  proRataCasualLeave: number;
  proRataSickLeave: number;
  totalAnnualLeaveAvailable: number;
  totalCasualLeaveAvailable: number;
  totalSickLeaveAvailable: number;
}

interface IApplyLeaveState {
  formState: IFormState;
  errors: IFormErrors;
  isSubmitting: boolean;
  showProbationPopup: boolean;
  showConfirmationPopup: boolean;
  weekendError: string | null;
  showHalfDayType: boolean;
  currentUser: any;
  leaveQuota: ILeaveQuota | null;
  employeeLeaveBalance: IEmployeeLeaveBalance | null;
  currentCycle: ILeaveCycle | null;
  usedLeavesInCycle: {
    AnnualLeaves: number;
    CasualLeaves: number;
    SickLeaves: number;
    OtherLeaves: number;
  };
  employee: IEmployeeData | null;
  loading: boolean;
  existingRequests: ILeaveRequest[] | null;
  proRataInfo: IProRataResult | null;
  showEditLeave: boolean; // Add this flag
}

const STATIC_LEAVE_TYPES: ILeaveType[] = [
  { Id: 1, Title: 'Annual Leave' },
  { Id: 2, Title: 'Casual Leave' },
  { Id: 3, Title: 'Sick Leave' },
  { Id: 4, Title: 'Other Leave' }
];

const FORM_INITIAL_STATE: IFormState = {
  leaveTypeId: 0,
  leaveDurationType: 'Full Day',
  startDate: '',
  endDate: '',
  totalDays: 0,
  halfDayType: 'First Half',
  reason: '',
  attachment: null,
  otherLeaveType: '',
};

const FILE_CONFIG = {
  MAX_SIZE: 5 * 1024 * 1024,
  ACCEPTED_TYPES: ['application/pdf', 'image/jpeg', 'image/png', 'image/jpg'] as const,
};

type AcceptedFileType = typeof FILE_CONFIG.ACCEPTED_TYPES[number];

const formatDate = (date: Date): string => date.toISOString().split('T')[0];
const isWeekend = (date: Date): boolean => date.getDay() === 5 || date.getDay() === 6;

const calculateTotalDays = (startDate: Date, endDate: Date, isHalfDay: boolean): number => {
  if (isHalfDay) return 0.5;
  let days = 0;
  const current = new Date(startDate);
  const end = new Date(endDate);
  while (current <= end) {
    days++;
    current.setDate(current.getDate() + 1);
  }
  return days;
};

const validateWeekendDates = (startDate: Date, endDate: Date): string | null => {
  if (isWeekend(startDate)) return 'Start date cannot be a weekend (Friday or Saturday)';
  if (isWeekend(endDate)) return 'End date cannot be a weekend (Friday or Saturday)';
  return null;
};

const isDateOverlap = (start1: Date, end1: Date, start2: Date, end2: Date): boolean =>
  start1 <= end2 && end1 >= start2;

const isEmployeeEligibleForLeaveQuota = (employmentType: string): boolean => {
  const eligibleTypes = ['Permanent', 'Contract', 'Part-time'];
  return eligibleTypes.includes(employmentType);
};

const calculateProRataForLeave = (
  cycleStartDate: Date,
  currentDate: Date,
  fullQuota: number
): number => {
  let monthsCompleted = 0;
  
  let monthDiff = (currentDate.getFullYear() - cycleStartDate.getFullYear()) * 12;
  monthDiff += currentDate.getMonth() - cycleStartDate.getMonth();
  
  if (currentDate.getDate() >= cycleStartDate.getDate()) {
    monthsCompleted = monthDiff + 1;
  } else {
    monthsCompleted = monthDiff;
  }
  
  monthsCompleted = Math.min(monthsCompleted, 12);
  monthsCompleted = Math.max(0, monthsCompleted);
  
  const proRataLeave = Math.round((monthsCompleted / 12) * fullQuota * 2) / 2;
  return proRataLeave;
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

export default class ApplyLeave extends React.Component<{ onLeaveSubmitted?: () => void }, IApplyLeaveState> {

  constructor(props: { onLeaveSubmitted?: () => void }) {
    super(props);
    this.state = {
      formState: FORM_INITIAL_STATE,
      errors: {},
      isSubmitting: false,
      showProbationPopup: false,
      showConfirmationPopup: false,
      weekendError: null,
      showHalfDayType: false,
      currentUser: null,
      leaveQuota: null,
      employeeLeaveBalance: null,
      currentCycle: null,
      usedLeavesInCycle: {
        AnnualLeaves: 0,
        CasualLeaves: 0,
        SickLeaves: 0,
        OtherLeaves: 0
      },
      employee: null,
      loading: true,
      existingRequests: null,
      proRataInfo: null,
      showEditLeave: false
    };
  }

  // Check URL parameters for RequestID
  checkUrlForRequestId = (): boolean => {
    const urlParams = new URLSearchParams(window.location.search);
    const requestIdParam = urlParams.get('RequestID');
    
    // Check if RequestID exists and has a valid numeric value
    if (requestIdParam && requestIdParam.trim() !== '') {
      const requestId = parseInt(requestIdParam, 10);
      if (!isNaN(requestId) && requestId > 0) {
        return true;
      }
    }
    return false;
  };

  async componentDidMount() {
    // Check if URL has valid RequestID
    const hasValidRequestId = this.checkUrlForRequestId();
    
    if (hasValidRequestId) {
      // Show EditLeave component
      this.setState({ showEditLeave: true, loading: false });
    } else {
      // Load normal form
      await this.fetchCurrentUser();
    }
  }

  fetchCurrentUser = async () => {
    try {
      const user = await sp.web.currentUser();
      this.setState({ currentUser: user }, () => {
        this.fetchEmployeeData();
      });
    } catch (err) {
      this.setState({ loading: false, errors: { submit: 'Failed to load user information' } });
    }
  };

  fetchEmployeeData = async () => {
    try {
      const { currentUser } = this.state;
      if (!currentUser?.Email) {
        throw new Error('User email not found');
      }

      const items = await sp.web.lists
        .getByTitle('Employee Information List')
        .items
        .select(
          'Id',
          'EmployeeName/Id',
          'EmployeeName/Title',
          'EmployeeName/EMail',
          'Department/Id',
          'Department/Title',
          'Country/Id',
          'Country/Title',
          'CountryCode/Id',
          'CountryCode/Title',
          'JoiningDate',
          'EmploymentType'
        )
        .expand('EmployeeName', 'Department', 'Country', 'CountryCode')
        .filter(`EmployeeName/EMail eq '${currentUser.Email.replace(/'/g, "''")}'`)
        .top(1)
        .get();

      if (!items || items.length === 0) {
        throw new Error('No employee record found');
      }

      this.processEmployeeData(items[0]);
    } catch (err) {
      this.setState({ loading: false, errors: { submit: 'Failed to load employee data' } });
    }
  };

  fetchDepartmentData = async (departmentId: number) => {
    try {
      if (!departmentId) return null;

      const items = await sp.web.lists
        .getByTitle('Department')
        .items
        .select('Id', 'Title', 'DepartmentManager/Id', 'DepartmentManager/Title', 'DepartmentManager/EMail')
        .expand('DepartmentManager')
        .filter(`Id eq ${departmentId}`)
        .top(1)
        .get();

      return items && items.length > 0 ? items[0] : null;
    } catch (err) {
      return null;
    }
  };

  fetchLeaveQuota = async (countryId: number, countryCodeId: number) => {
    try {
      if (!countryId || !countryCodeId) {
        return null;
      }

      const items = await sp.web.lists
        .getByTitle('Leave Quota')
        .items
        .select(
          'Id',
          'Country/Id',
          'CountryCode/Id',
          'Leaves',
          'AnnualLeaves',
          'CasualLeaves',
          'SickLeaves',
          'OtherLeaves'
        )
        .expand('Country', 'CountryCode')
        .filter(`Country/Id eq ${countryId} and CountryCode/Id eq ${countryCodeId}`)
        .top(1)
        .get();

      if (items && items.length > 0) {
        const quota = items[0];
        return {
          Id: quota.Id,
          CountryId: quota.Country?.Id,
          CountryCodeId: quota.CountryCode?.Id,
          Leaves: quota.Leaves,
          AnnualLeaves: quota.AnnualLeaves,
          CasualLeaves: quota.CasualLeaves,
          SickLeaves: quota.SickLeaves,
          OtherLeaves: quota.OtherLeaves
        } as ILeaveQuota;
      }
      return null;
    } catch (err) {
      return null;
    }
  };

  fetchLeaveRequests = async (employeeId: number, cycleStart: Date, cycleEnd: Date) => {
    try {
      if (!employeeId) return [];

      const startDateStr = formatDate(cycleStart);
      const endDateStr = formatDate(cycleEnd);

      const filter = `EmployeeId eq ${employeeId} and (Status eq 'Pending on Manager' or Status eq 'Pending on HR' or Status eq 'Approved') and StartDate ge '${startDateStr}' and EndDate le '${endDateStr}'`;

      const items = await sp.web.lists
        .getByTitle('Leave Requests')
        .items
        .select('Id', 'StartDate', 'EndDate', 'Status', 'TotalDays', 'LeaveTypeId')
        .filter(filter)
        .orderBy('StartDate', false)
        .get();

      return items as ILeaveRequest[];
    } catch (err) {
      return [];
    }
  };

  fetchLeaveBalance = async (employeeId: number) => {
    try {
      if (!employeeId) return null;

      const currentYear = new Date().getFullYear().toString();

      const items = await sp.web.lists
        .getByTitle('Employee Leave Balance')
        .items
        .select('Id', 'LeavesBalance', 'Used', 'Remaining', 'Year')
        .filter(`EmployeeId eq ${employeeId} and Year eq '${currentYear}'`)
        .top(1)
        .get();

      return items && items.length > 0 ? items[0] as IEmployeeLeaveBalance : null;
    } catch (err) {
      return null;
    }
  };

  processEmployeeData = async (empData: any) => {
    const { currentUser } = this.state;

    const departmentId = empData.Department?.Id;
    const deptData = departmentId ? await this.fetchDepartmentData(departmentId) : null;

    let managerId = 0, managerName = '', managerEmail = '';
    if (deptData?.DepartmentManager) {
      managerId = deptData.DepartmentManager.Id;
      managerName = deptData.DepartmentManager.Title;
      managerEmail = deptData.DepartmentManager.EMail;
    }

    const employee: IEmployeeData = {
      id: currentUser?.Id || 0,
      name: empData.EmployeeName?.Title || currentUser?.Title || '',
      department: empData.Department?.Title || '',
      departmentId: departmentId || 0,
      country: empData.Country?.Title || '',
      countryId: empData.Country?.Id || 0,
      countryCode: empData.CountryCode?.Title || '',
      countryCodeId: empData.CountryCode?.Id || 0,
      managerId, managerName, managerEmail,
      joinDate: empData.JoiningDate || '',
      employmentType: empData.EmploymentType || 'Permanent',
      userImageUrl: '',
    };

    this.setState({ employee }, async () => {
      const cycle = calculateCurrentLeaveCycle(employee.joinDate);
      this.setState({ currentCycle: cycle });

      await Promise.all([
        this.fetchLeaveQuota(employee.countryId, employee.countryCodeId),
        this.fetchLeaveBalance(employee.id),
        this.fetchLeaveRequests(employee.id, cycle.startDate, cycle.endDate)
      ]).then(([quota, balance, requests]) => {
        this.setState({
          leaveQuota: quota,
          employeeLeaveBalance: balance,
          existingRequests: requests,
          loading: false
        }, () => {
          this.calculateUsedLeaves();
          this.checkProbationStatus();
          this.calculateProRataLeave();
        });
      });
    });
  };

  calculateProRataLeave = () => {
    const { employee, leaveQuota, employeeLeaveBalance, currentCycle, usedLeavesInCycle } = this.state;
    
    if (!employee || !leaveQuota || !currentCycle || this.isProbation()) {
      this.setState({ proRataInfo: null });
      return;
    }

    if (!isEmployeeEligibleForLeaveQuota(employee.employmentType)) {
      this.setState({ proRataInfo: null });
      return;
    }

    const currentDate = new Date();
    const cycleStartDate = currentCycle.startDate;
    const cycleNumber = currentCycle.cycleNumber;
    
    const annualQuota = leaveQuota.AnnualLeaves;
    const casualQuota = leaveQuota.CasualLeaves;
    const sickQuota = leaveQuota.SickLeaves;
    
    const proRataAnnual = calculateProRataForLeave(cycleStartDate, currentDate, annualQuota);
    const proRataCasual = calculateProRataForLeave(cycleStartDate, currentDate, casualQuota);
    const proRataSick = calculateProRataForLeave(cycleStartDate, currentDate, sickQuota);
    
    let totalAnnual = proRataAnnual;
    let totalCasual = proRataCasual;
    let totalSick = proRataSick;
    
    if (cycleNumber > 1 && employeeLeaveBalance) {
      totalAnnual = proRataAnnual + (employeeLeaveBalance.Remaining || 0);
    }
    
    totalAnnual = Math.max(0, totalAnnual - usedLeavesInCycle.AnnualLeaves);
    totalCasual = Math.max(0, totalCasual - usedLeavesInCycle.CasualLeaves);
    totalSick = Math.max(0, totalSick - usedLeavesInCycle.SickLeaves);
    
    this.setState({
      proRataInfo: {
        proRataAnnualLeave: proRataAnnual,
        proRataCasualLeave: proRataCasual,
        proRataSickLeave: proRataSick,
        totalAnnualLeaveAvailable: totalAnnual,
        totalCasualLeaveAvailable: totalCasual,
        totalSickLeaveAvailable: totalSick
      }
    });
  };

  checkProbationStatus = () => {
    const { employee } = this.state;
    if (!employee?.employmentType) return;

    const isProbation = employee.employmentType.toLowerCase() === 'probation';

    if (isProbation) {
      const otherLeaveType = STATIC_LEAVE_TYPES.find(type => type.Title === 'Other Leave');
      if (otherLeaveType) {
        this.handleFieldChange('leaveTypeId', otherLeaveType.Id);
      }
      this.setState({ showProbationPopup: true });
    }
  };

  calculateUsedLeaves = () => {
    const { existingRequests, currentCycle } = this.state;
    if (existingRequests && existingRequests.length > 0 && currentCycle) {
      const used = { AnnualLeaves: 0, CasualLeaves: 0, SickLeaves: 0, OtherLeaves: 0 };

      existingRequests.forEach((request: ILeaveRequest) => {
        const totalDays = request.TotalDays || 0;
        if (request.LeaveTypeId === 1) used.AnnualLeaves += totalDays;
        else if (request.LeaveTypeId === 2) used.CasualLeaves += totalDays;
        else if (request.LeaveTypeId === 3) used.SickLeaves += totalDays;
        else if (request.LeaveTypeId === 4) used.OtherLeaves += totalDays;
      });

      this.setState({ usedLeavesInCycle: used }, () => {
        this.calculateProRataLeave();
      });
    }
  };

  isProbation = (): boolean => {
    const { employee } = this.state;
    if (!employee?.employmentType) return false;
    return employee.employmentType.toLowerCase() === 'probation';
  };

  getAvailableQuota = (leaveTypeId: number): number => {
    const { leaveQuota, usedLeavesInCycle, proRataInfo } = this.state;

    if (this.isProbation()) return 0;
    if (!isEmployeeEligibleForLeaveQuota(this.state.employee?.employmentType || '')) return 0;
    if (!leaveQuota) return 0;

    switch (leaveTypeId) {
      case 1:
        if (proRataInfo) {
          return proRataInfo.totalAnnualLeaveAvailable;
        }
        return Math.max(0, leaveQuota.AnnualLeaves - usedLeavesInCycle.AnnualLeaves);
      
      case 2:
        if (proRataInfo) {
          return proRataInfo.totalCasualLeaveAvailable;
        }
        return Math.max(0, leaveQuota.CasualLeaves - usedLeavesInCycle.CasualLeaves);
      
      case 3:
        if (proRataInfo) {
          return proRataInfo.totalSickLeaveAvailable;
        }
        return Math.max(0, leaveQuota.SickLeaves - usedLeavesInCycle.SickLeaves);
      
      case 4:
        return Math.max(0, leaveQuota.OtherLeaves - usedLeavesInCycle.OtherLeaves);
      
      default:
        return 0;
    }
  };

  canApplyForLeave = (leaveTypeId: number, requestedDays: number): boolean => {
    if (leaveTypeId === 4) return true;
    if (this.isProbation()) return false;
    if (!isEmployeeEligibleForLeaveQuota(this.state.employee?.employmentType || '')) return false;

    const available = this.getAvailableQuota(leaveTypeId);
    return requestedDays <= available;
  };

  validateForm = (): IFormErrors => {
    const { formState, existingRequests } = this.state;
    const newErrors: IFormErrors = {};
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    if (!formState.leaveTypeId || formState.leaveTypeId === 0) newErrors.leaveType = 'Please select a leave type';
    if (formState.leaveTypeId === 4 && !formState.otherLeaveType.trim()) newErrors.otherLeaveType = 'Please specify other leave type';

    if (!formState.startDate) {
      newErrors.startDate = 'Please select start date';
    } else {
      const startDate = new Date(formState.startDate);
      if (startDate < today) newErrors.startDate = 'Cannot apply for past dates';
      if (isWeekend(startDate)) newErrors.startDate = 'Start date cannot be a weekend (Friday or Saturday)';
    }

    if (formState.leaveDurationType === 'Full Day' && !formState.endDate) {
      newErrors.endDate = 'Please select end date';
    } else if (formState.startDate && formState.endDate) {
      const start = new Date(formState.startDate);
      const end = new Date(formState.endDate);
      if (end < start) newErrors.endDate = 'End date cannot be before start date';
      if (isWeekend(end)) newErrors.endDate = 'End date cannot be a weekend (Friday or Saturday)';
    }

    if (!formState.reason) {
      newErrors.reason = 'Please provide a reason for leave';
    } else if (formState.reason.length < 10) {
      newErrors.reason = 'Please provide a more detailed reason (minimum 10 characters)';
    }

    if (existingRequests && existingRequests.length > 0 && formState.startDate && formState.endDate) {
      const newStart = new Date(formState.startDate);
      const newEnd = new Date(formState.endDate);
      const hasOverlap = existingRequests.some((leave: ILeaveRequest) => {
        const existingStart = new Date(leave.StartDate);
        const existingEnd = new Date(leave.EndDate);
        return isDateOverlap(newStart, newEnd, existingStart, existingEnd);
      });
      if (hasOverlap) newErrors.overlap = 'You already have a leave request for these dates';
    }

    if (formState.leaveTypeId && formState.leaveTypeId !== 4 && formState.totalDays > 0) {
      const canApply = this.canApplyForLeave(formState.leaveTypeId, formState.totalDays);
      const selectedLeaveType = STATIC_LEAVE_TYPES.find(type => type.Id === formState.leaveTypeId);
      const availableQuota = this.getAvailableQuota(formState.leaveTypeId);

      if (!canApply) {
        if (availableQuota <= 0) {
          newErrors.insufficientQuota = `You have no ${selectedLeaveType?.Title} quota remaining.`;
        } else {
          newErrors.insufficientQuota = `Insufficient ${selectedLeaveType?.Title} quota. Available: ${availableQuota} days, Requested: ${formState.totalDays} days.`;
        }
      }
    }

    const selectedLeaveType = STATIC_LEAVE_TYPES.find(type => type.Id === formState.leaveTypeId);
    if (selectedLeaveType?.Title === 'Sick Leave' && formState.totalDays > 1 && !formState.attachment) {
      newErrors.attachment = 'Medical certificate required for sick leave exceeding 1 day';
    }

    if (formState.attachment) {
      const isValidType = FILE_CONFIG.ACCEPTED_TYPES.includes(formState.attachment.type as AcceptedFileType);
      if (!isValidType) newErrors.attachment = 'Only PDF, JPG, and PNG files are allowed';
      else if (formState.attachment.size > FILE_CONFIG.MAX_SIZE) newErrors.attachment = 'File size must be less than 5MB';
    }

    return newErrors;
  };

  handleFieldChange = (field: keyof IFormState, value: any) => {
    this.setState(prevState => ({
      formState: { ...prevState.formState, [field]: value },
      errors: { ...prevState.errors, [field]: undefined }
    }), () => {
      if (field === 'leaveDurationType') {
        const isHalfDay = value === 'Half Day';
        this.setState({ showHalfDayType: isHalfDay });

        if (isHalfDay && this.state.formState.startDate) {
          this.setState(prevState => ({
            formState: { ...prevState.formState, endDate: prevState.formState.startDate }
          }), () => {
            this.updateTotalDays();
          });
        } else if (!isHalfDay) {
          const { startDate, endDate } = this.state.formState;
          if (startDate && endDate && new Date(startDate) > new Date(endDate)) {
            this.setState(prevState => ({
              formState: { ...prevState.formState, endDate: '' }
            }), () => {
              this.updateTotalDays();
            });
          } else {
            this.updateTotalDays();
          }
        }
      } else if (field === 'startDate') {
        const isHalfDay = this.state.formState.leaveDurationType === 'Half Day';

        if (isHalfDay) {
          this.setState(prevState => ({
            formState: { ...prevState.formState, endDate: value }
          }), () => {
            this.updateTotalDays();
          });
        } else {
          const { endDate } = this.state.formState;
          if (value && endDate && new Date(value) > new Date(endDate)) {
            this.setState(prevState => ({
              formState: { ...prevState.formState, endDate: '' }
            }), () => {
              this.updateTotalDays();
            });
          } else {
            this.updateTotalDays();
          }
        }
      } else if (field === 'endDate') {
        this.updateTotalDays();
      }
    });
  };

  handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      this.handleFieldChange('attachment', e.target.files[0]);
    }
  };

  updateTotalDays = () => {
    const { startDate, endDate, leaveDurationType } = this.state.formState;

    if (leaveDurationType === 'Half Day' && startDate) {
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: 0.5 }
      }));
      return;
    }

    if (!startDate || !endDate) {
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: 0 }
      }));
      return;
    }

    const start = new Date(startDate);
    const end = new Date(endDate);
    if (isNaN(start.getTime()) || isNaN(end.getTime())) return;

    const weekendErrorMsg = validateWeekendDates(start, end);
    this.setState({ weekendError: weekendErrorMsg });

    if (!weekendErrorMsg && end >= start) {
      const days = calculateTotalDays(start, end, false);
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: days }
      }));
    } else {
      this.setState(prevState => ({
        formState: { ...prevState.formState, totalDays: 0 }
      }));
    }
  };

  handleSubmit = () => {
    const { formState } = this.state;

    if (formState.startDate && formState.endDate) {
      const start = new Date(formState.startDate);
      const end = new Date(formState.endDate);
      const weekendErrorMsg = validateWeekendDates(start, end);
      if (weekendErrorMsg) {
        this.setState({ errors: { weekend: weekendErrorMsg } });
        return;
      }
    }

    const validationErrors = this.validateForm();
    if (Object.keys(validationErrors).length > 0) {
      this.setState({ errors: validationErrors });
      return;
    }

    if (!this.state.employee?.managerId) {
      this.setState({ errors: { manager: 'No manager assigned. Please contact HR.' } });
      return;
    }

    this.setState({ showConfirmationPopup: true });
  };

  confirmSubmit = async () => {
    const { formState, employee } = this.state;
    const { onLeaveSubmitted } = this.props;

    this.setState({ isSubmitting: true, showConfirmationPopup: false });

    try {
      const leaveRequestData: any = {
        Title: `Leave Request - ${employee?.name}`,
        EmployeeId: employee?.id,
        LineManagerId: employee?.managerId,
        DepartmentId: employee?.departmentId,
        CountryId: employee?.countryId,
        CountryCodeId: employee?.countryCodeId,
        LeaveTypeId: formState.leaveTypeId,
        OtherLeaveType: formState.otherLeaveType || null,
        LeaveDurationType: formState.leaveDurationType,
        HalfDayType: formState.halfDayType,
        StartDate: formState.startDate,
        EndDate: formState.endDate,
        TotalDays: formState.totalDays,
        Reason: formState.reason,
        Status: 'Pending on Manager',
        EmploymentType: employee?.employmentType,
        EmployeeYearCycle: this.state.currentCycle ? this.state.currentCycle.cycleNumber : 404
      };

      const result = await sp.web.lists.getByTitle('Leave Requests').items.add(leaveRequestData);

      if (formState.attachment) {
        const buffer = await formState.attachment.arrayBuffer();
        await sp.web.lists.getByTitle('Leave Requests')
          .items.getById(result.data.Id)
          .attachmentFiles.add(formState.attachment.name, buffer);
      }

      this.setState({
        formState: { ...FORM_INITIAL_STATE, leaveTypeId: formState.leaveTypeId === 4 ? 4 : 0 },
        errors: {},
        weekendError: null
      });

      if (employee && this.state.currentCycle) {
        const updatedRequests = await this.fetchLeaveRequests(employee.id, this.state.currentCycle.startDate, this.state.currentCycle.endDate);
        this.setState({ existingRequests: updatedRequests }, () => {
          this.calculateUsedLeaves();
        });
      }

      if (onLeaveSubmitted) onLeaveSubmitted();
      alert('Leave request submitted successfully!');
    } catch (error: any) {
      this.setState({ errors: { submit: error.message || 'Error submitting leave request' } });
    } finally {
      this.setState({ isSubmitting: false });
    }
  };

  componentDidUpdate(prevProps: any, prevState: IApplyLeaveState) {
    if (prevState.formState.leaveDurationType !== this.state.formState.leaveDurationType &&
      this.state.formState.leaveDurationType === 'Half Day' &&
      this.state.formState.startDate) {
      this.handleFieldChange('endDate', this.state.formState.startDate);
    }
  }

  render() {
    const { showEditLeave, loading, employee, formState, errors, isSubmitting, showProbationPopup, showConfirmationPopup, weekendError, showHalfDayType } = this.state;

    // Agar URL mein valid RequestID hai to EditLeave component dikhao
    if (showEditLeave) {
      return <EditLeave />;
    }

    if (loading) {
      return (
        <div className={styles.loadingContainer}>
          <div className={styles.spinner}></div>
          <p>Loading your leave information...</p>
        </div>
      );
    }

    if (!employee) {
      return (
        <div className={styles.errorContainer}>
          <p>No employee data found. Please contact HR.</p>
          <button onClick={() => window.location.reload()}>Retry</button>
        </div>
      );
    }

    return (
      <>
        {showProbationPopup && (
          <div className={styles.probationPopupOverlay} onClick={() => this.setState({ showProbationPopup: false })}>
            <div className={styles.probationPopup} onClick={e => e.stopPropagation()}>
              <div className={styles.probationPopupHeader}>
                <h3>Probation Period Notice</h3>
              </div>
              <div className={styles.probationPopupBody}>
                <p>Dear <strong>{employee.name}</strong>,</p>
                <p>As per company policy, employees on <strong>Probation Period</strong> are only eligible to apply for <strong>Other Leave</strong> during their probation period.</p>
                <p>All other leave types will be available after successful completion of probation.</p>
              </div>
              <div className={styles.probationPopupFooter}>
                <button className={styles.probationPopupButton} onClick={() => this.setState({ showProbationPopup: false })}>Got It</button>
              </div>
            </div>
          </div>
        )}

        {showConfirmationPopup && (
          <div className={styles.probationPopupOverlay} onClick={() => this.setState({ showConfirmationPopup: false })}>
            <div className={styles.probationPopup} onClick={e => e.stopPropagation()}>
              <div className={styles.probationPopupHeader}>
                <h3>Confirm Leave Request</h3>
              </div>
              <div className={styles.probationPopupBody}>
                <p>Are you sure you want to submit this leave request?</p>
                <p><strong>Leave Type:</strong> {STATIC_LEAVE_TYPES.find(t => t.Id === formState.leaveTypeId)?.Title}</p>
                <p><strong>Duration:</strong> {formState.totalDays} day(s)</p>
                <p><strong>Dates:</strong> {formState.startDate} to {formState.endDate}</p>
                <p><strong>Reason:</strong> {formState.reason}</p>
              </div>
              <div className={styles.probationPopupFooter} style={{ gap: '12px' }}>
                <button className={styles.probationPopupButton} style={{ background: '#94a3b8' }} onClick={() => this.setState({ showConfirmationPopup: false })}>Cancel</button>
                <button className={styles.probationPopupButton} onClick={this.confirmSubmit} disabled={isSubmitting}>
                  {isSubmitting ? 'Submitting...' : 'Confirm'}
                </button>
              </div>
            </div>
          </div>
        )}

        <div className={styles.applyLeaveFormContent}>
          {(Object.values(errors).some(error => error) || weekendError) && (
            <div className={styles.errorSummary}>
              {weekendError && <div className={styles.errorMessage}>{weekendError}</div>}
              {Object.values(errors).map((error, idx) => error && <div key={idx} className={styles.errorMessage}>{error}</div>)}
            </div>
          )}

          <div className={styles.applyLeaveForm}>
            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Employee Name</label>
                <input type="text" value={employee.name} disabled className={styles.applyLeaveDisabledField} />
              </div>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Department</label>
                <input type="text" value={employee.department} disabled className={styles.applyLeaveDisabledField} />
              </div>
            </div>

            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Country</label>
                <input type="text" value={employee.country} disabled className={styles.applyLeaveDisabledField} />
              </div>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Manager</label>
                <input type="text" value={employee.managerName || 'Not Assigned'} disabled className={styles.applyLeaveDisabledField} />
              </div>
            </div>

            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Leave Type <span className={styles.applyLeaveRequired}>*</span></label>
                <select value={formState.leaveTypeId} onChange={e => this.handleFieldChange('leaveTypeId', parseInt(e.target.value))} className={styles.applyLeaveSelect} disabled={this.isProbation()}>
                  <option value="0">Select Leave Type</option>
                  {STATIC_LEAVE_TYPES.map(type => (
                    <option key={type.Id} value={type.Id}>
                      {type.Title}
                    </option>
                  ))}
                </select>
                {this.isProbation() && <div className={styles.probationNote}>* During probation period, only "Other Leave" can be selected</div>}
              </div>

              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Leave Duration <span className={styles.applyLeaveRequired}>*</span></label>
                <select value={formState.leaveDurationType} onChange={e => this.handleFieldChange('leaveDurationType', e.target.value as 'Full Day' | 'Half Day')} className={styles.applyLeaveSelect}>
                  <option value="Full Day">Full Day</option>
                  <option value="Half Day">Half Day</option>
                </select>
              </div>
            </div>

            {formState.leaveTypeId === 4 && (
              <div className={styles.applyLeaveRow}>
                <div className={styles.applyLeaveFormGroup}>
                  <label className={styles.applyLeaveLabel}>Other Leave Type <span className={styles.applyLeaveRequired}>*</span></label>
                  <input type="text" value={formState.otherLeaveType} onChange={e => this.handleFieldChange('otherLeaveType', e.target.value)} placeholder="e.g., Emergency Leave, Marriage Leave, etc." className={styles.applyLeaveInput} />
                </div>
                {showHalfDayType && (
                  <div className={styles.applyLeaveFormGroup}>
                    <label className={styles.applyLeaveLabel}>Half Day Type <span className={styles.applyLeaveRequired}>*</span></label>
                    <select value={formState.halfDayType} onChange={e => this.handleFieldChange('halfDayType', e.target.value as 'First Half' | 'Second Half')} className={styles.applyLeaveSelect}>
                      <option value="First Half">First Half</option>
                      <option value="Second Half">Second Half</option>
                    </select>
                  </div>
                )}
              </div>
            )}

            {formState.leaveTypeId !== 4 && showHalfDayType && (
              <div className={styles.applyLeaveRow}>
                <div className={styles.applyLeaveFormGroup}>
                  <label className={styles.applyLeaveLabel}>Half Day Type <span className={styles.applyLeaveRequired}>*</span></label>
                  <select value={formState.halfDayType} onChange={e => this.handleFieldChange('halfDayType', e.target.value as 'First Half' | 'Second Half')} className={styles.applyLeaveSelect}>
                    <option value="First Half">First Half</option>
                    <option value="Second Half">Second Half</option>
                  </select>
                </div>
              </div>
            )}

            <div className={styles.applyLeaveRow}>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>Start Date <span className={styles.applyLeaveRequired}>*</span></label>
                <input type="date" value={formState.startDate} min={formatDate(new Date())} onChange={e => this.handleFieldChange('startDate', e.target.value)} className={styles.applyLeaveInput} />
              </div>
              <div className={styles.applyLeaveFormGroup}>
                <label className={styles.applyLeaveLabel}>End Date {formState.leaveDurationType !== 'Half Day' && <span className={styles.applyLeaveRequired}>*</span>}</label>
                <input type="date" value={formState.endDate} disabled={formState.leaveDurationType === 'Half Day'} min={formState.startDate || formatDate(new Date())} onChange={e => this.handleFieldChange('endDate', e.target.value)} className={styles.applyLeaveInput} />
              </div>
            </div>

            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Total Days</label>
              <input type="number" value={formState.totalDays} disabled className={styles.applyLeaveTotalDaysField} step="0.5" />
            </div>

            <div className={styles.applyLeaveFormGroup}>
              <label className={styles.applyLeaveLabel}>Reason for Leave <span className={styles.applyLeaveRequired}>*</span></label>
              <textarea value={formState.reason} onChange={e => this.handleFieldChange('reason', e.target.value)} placeholder="Please provide reason for your leave" rows={4} className={styles.applyLeaveTextarea} />
            </div>

            {(() => {
              const selectedLeaveType = STATIC_LEAVE_TYPES.find(type => type.Id === formState.leaveTypeId);
              return selectedLeaveType?.Title === 'Sick Leave' && formState.totalDays > 1 && (
                <div className={styles.applyLeaveFormGroup}>
                  <label className={styles.applyLeaveLabel}>Medical Certificate <span className={styles.applyLeaveRequired}>*</span></label>
                  <div className={styles.applyLeaveFileDropZone} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); if (e.dataTransfer.files?.[0]) this.handleFieldChange('attachment', e.dataTransfer.files[0]); }} onClick={() => document.getElementById('applyLeaveFileInput')?.click()}>
                    {formState.attachment ? (
                      <div className={styles.applyLeaveFilePreview}>
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M21 15V19C21 19.5304 20.7893 20.0391 20.4142 20.4142C20.0391 20.7893 19.5304 21 19 21H5C4.46957 21 3.96086 20.7893 3.58579 20.4142C3.21071 20.0391 3 19.5304 3 19V15M7 10L12 15M12 15L17 10M12 15V3" stroke="currentColor" strokeWidth="1.5" /></svg>
                        <span>{formState.attachment.name}</span>
                        <span className={styles.applyLeaveRemoveFile} onClick={e => { e.stopPropagation(); this.handleFieldChange('attachment', null); }}><svg width="14" height="14" viewBox="0 0 24 24" fill="none"><path d="M18 6L6 18M6 6L18 18" stroke="currentColor" strokeWidth="2" /></svg></span>
                      </div>
                    ) : (
                      <div className={styles.applyLeaveUploadPrompt}>
                        <svg width="32" height="32" viewBox="0 0 24 24" fill="none"><path d="M12 16V4M12 4L8 8M12 4L16 8M20 16V19C20 19.5304 19.7893 20.0391 19.4142 20.4142C19.0391 20.7893 18.5304 21 18 21H6C4.46957 21 3.96086 20.7893 3.58579 20.4142C3.21071 20.0391 3 19.5304 3 19V16" stroke="currentColor" strokeWidth="1.5" /></svg>
                        <span>Click to upload or drag & drop</span>
                        <span className={styles.applyLeaveFileHint}>Supported: PDF, JPG, PNG (Max 5MB)</span>
                      </div>
                    )}
                  </div>
                  <input id="applyLeaveFileInput" type="file" style={{ display: 'none' }} accept=".pdf,.jpg,.jpeg,.png" onChange={this.handleFileChange} />
                </div>
              );
            })()}

            <div className={styles.applyLeaveButtonContainer}>
              <button type="button" className={styles.applyLeaveSubmitButton} onClick={this.handleSubmit} disabled={isSubmitting}>
                {isSubmitting ? 'Submitting...' : 'Submit Leave Request'}
              </button>
            </div>
          </div>
        </div>
      </>
    );
  }
}