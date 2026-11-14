<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Hash;
use Illuminate\Support\Facades\Validator;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Str;

use Mail;

// MODELS
use App\Model\AuditResult;
use App\Model\ARSection;
use App\Model\ARProductLine;
use App\Model\ARAuditor;
use App\Model\ARAuditee;
use App\Model\AREmailRecipient;
use App\Exports\AuditResultsExport; // Update by JD
use App\Model\FactorItemList; // Update by JD

// PACKAGES
use DataTables;
use Maatwebsite\Excel\Facades\Excel;

class AuditResultController extends Controller
{
    public function view_audit_results_test(Request $request){
        session_start();
        $audit_results;
        $userId = 82;
        // if(isset($request->status)){
            $audit_results = AuditResult::with([
                    'checked_by_info',
                    // 'product_line_details.product_line_info',
                    // 'section_details.section_info',
                    'auditor_details.user_info',
                    'auditee_details',
                ])
                // ->orDoesntHave('checked_by_info')
                ->whereHas('auditor_details', function($query) use($userId){
                    $query->where('user_id', 190);
                })
                ->where('layer', 1)
                ->orWhereHas('auditee_details', function($query) use($userId){
                    $query->where('user_id', 190);
                })
                ->where('layer', 1)
                ->orWhere('checked_by', 190)
                ->where('layer', 1)
                // ->where('layer', 1)
                ->take(2)
                ->get();

            return $audit_results;
                // return collect($audit_results)->where('checked_by', '!=', null);
                // return collect($audit_results);
    }

    public function view_audit_results(Request $request){
        session_start();
        $audit_results;
        $userId = $request->userId;
        // if(isset($request->status)){
            $audit_results = AuditResult::with([
                'checked_by_info',
                'product_line_details.product_line_info',
                'section_details.section_info',
                'auditor_details.user_info', // In this case, I filtered the relation first, before using the whereHas() method to be able to target the auditor_details(relation method) as you can see below, auditor_details(relation method) has user_id(column) inside it and use it for filter
                'auditee_details.user_info',
            ])

            ->whereHas('auditor_details', function($query) use($userId){
                $query->where('user_id', $userId);
            })
            ->where('layer', $request->layer)
            ->orWhereHas('auditee_details', function($query) use($userId){
                $query->where('user_id', $userId);
            })
            // ->has('checked_by_info', '>', 0)
            ->where('layer', $request->layer) // In this case, I use 2 where() since I use the orWhereHas() because if orWhereHas is true, it will fetched the data layer 2 & layer 3 even if I specified the layer = 1. Basically this means to be specific, when everytime I use the eloquent syntax that starts with "or" then repeat the code(query) that should be matched, like on the query above, the layer should be the basis of all the data.
            ->orWhere('checked_by', $userId)
            ->where('layer', $request->layer)
            ->orderBy('id', 'desc')
            /**
             * The Queries of the above code
             * select * from `audit_results` where exists (select * from `ar_auditors` where `audit_results`.`id` = `ar_auditors`.`audit_result_id` and `user_id` = 190) and `layer` = '1'
             * or exists (select * from `ar_auditees` where `audit_results`.`id` = `ar_auditees`.`audit_result_id` and `user_id` = 190) and `layer` = '1'
             */
            ->get();
        // }
        // else {
        //     $audit_results = AuditResult::where('status', $request->status)
        //     ->where('layer', $request->layer)
        //     ->get();
        // }
            // return $request->userId;
        return DataTables::of($audit_results)
            ->addColumn('raw_status', function($data){
                $result = "";

                if($data->status == 1){
                    $result .= '<span class="tag tag-success">Active</span>';
                }
                else{
                    $result .= '<span class="tag tag-danger">Inactive</span>';
                }

                return $result;
            })
            ->addColumn('raw_audit_status', function($data){
                $result = "";

                if($data->audit_status == 1){
                    $result .= '<span class="tag tag-warning">For Checking</span>';
                }
                else if($data->audit_status == 2){
                    $result .= '<span class="tag tag-primary">For CAPA</span>';
                }
                else if($data->audit_status == 3){
                    $result .= '<span class="tag tag-info">For Close Out</span>';
                }
                else if($data->audit_status == 4){
                    $result .= '<span class="tag tag-success">Closed</span>';
                }
                else{
                    $result .= '<span class="tag tag-info">No Findings</span>';
                }

                return $result;
            })
            ->addColumn('raw_action', function($data){
                $result = '';
                if($data->status == 1) {
                    $result .= '<button class="btn-actions btn btn-sm btn-success btnEditAuditResult" audit_result-id="' . $data->id . '" type="button" title="Edit" layer="' . $data->layer . '"><i class="icon-edit2"></i></button> ';

                    $result .= '<button class="btn-actions btn btn-sm btn-danger btnActionAuditResult" type="button" audit_result-id="' . $data->id . '" status="2" title="Archive"><i class="icon-trash2"></i></button> ';

                    // $result .= '<button class="btn-actions btn btn-sm btn-info btnSelectAuditResult" type="button" audit_result-id="' . $data->id . '" audit_result-name="' . $data->name . '" title="Select"><i class="icon-arrow-right"></i></button> ';
                }
                else {
                    $result .= '<button class="btn-actions btn btn-sm btn-warning btnActionAuditResult" type="button" audit_result-id="' . $data->id . '" status="1" title="Restore"><i class="icon-history2"></i></button> ';
                }

                return $result;
            })
            ->addColumn('raw_department', function($data){
                $result = "";

                if($data->department == 1){
                    $result .= 'TS';
                }
                else if($data->department == 2){
                    $result .= 'PPS';
                }
                else if($data->department == 3){
                    $result .= 'CN';
                }
                else{
                    $result .= 'YF';
                }

                return $result;
            })
            ->rawColumns(['raw_status', 'raw_action', 'raw_department', 'raw_audit_status'])
            ->make(true);
    }

    // public function view_audit_results_layer_1(Request $request){
    //     session_start();
    //     $audit_results;
    //     $userId = $request->userId;
    //     // if(isset($request->status)){
    //         $audit_results = AuditResult::with([
    //                 'checked_by_info',
    //                 'product_line_details.product_line_info',
    //                 'section_details.section_info',
    //                 'auditor_details.user_info',
    //                 'auditee_details.user_info',
    //             ])
    //             ->where('layer', $request->layer)
    //             ->whereHas('auditor_details', function($query) use($userId){
    //                 $query->where('user_id', $userId);
    //             })
    //             ->orWhereHas('auditee_details', function($query) use($userId){
    //                 $query->where('user_id', $userId);
    //             })
    //             // ->whereHas('auditee_details', function($query){
    //             //     $query->whereIn('user_id', [190]);
    //             // })
    //             ->get();
    //     // }
    //     // else {
    //     //     $audit_results = AuditResult::where('status', $request->status)
    //     //     ->where('layer', $request->layer)
    //     //     ->get();
    //     // }
    //         // return $request->userId;
    //     return DataTables::of($audit_results)
    //         ->addColumn('raw_status', function($data){
    //             $result = "";

    //             if($data->status == 1){
    //                 $result .= '<span class="tag tag-success">Active</span>';
    //             }
    //             else{
    //                 $result .= '<span class="tag tag-danger">Inactive</span>';
    //             }

    //             return $result;
    //         })
    //         ->addColumn('raw_audit_status', function($data){
    //             $result = "";

    //             if($data->audit_status == 1){
    //                 $result .= '<span class="tag tag-warning">For Checking</span>';
    //             }
    //             else if($data->audit_status == 2){
    //                 $result .= '<span class="tag tag-primary">For CAPA</span>';
    //             }
    //             else if($data->audit_status == 3){
    //                 $result .= '<span class="tag tag-info">For Close Out</span>';
    //             }
    //             else if($data->audit_status == 4){
    //                 $result .= '<span class="tag tag-success">Closed</span>';
    //             }
    //             else{
    //                 $result .= '<span class="tag tag-danger">Unclosed</span>';
    //             }

    //             return $result;
    //         })
    //         ->addColumn('raw_action', function($data){
    //             $result = '';
    //             if($data->status == 1) {
    //                 $result .= '<button class="btn-actions btn btn-sm btn-success btnEditAuditResult" audit_result-id="' . $data->id . '" type="button" title="Edit" layer="' . $data->layer . '"><i class="icon-edit2"></i></button> ';

    //                 $result .= '<button class="btn-actions btn btn-sm btn-danger btnActionAuditResult" type="button" audit_result-id="' . $data->id . '" status="2" title="Archive"><i class="icon-trash2"></i></button> ';

    //                 // $result .= '<button class="btn-actions btn btn-sm btn-info btnSelectAuditResult" type="button" audit_result-id="' . $data->id . '" audit_result-name="' . $data->name . '" title="Select"><i class="icon-arrow-right"></i></button> ';
    //             }
    //             else {
    //                 $result .= '<button class="btn-actions btn btn-sm btn-warning btnActionAuditResult" type="button" audit_result-id="' . $data->id . '" status="1" title="Restore"><i class="icon-history2"></i></button> ';
    //             }

    //             return $result;
    //         })
    //         ->addColumn('raw_department', function($data){
    //             $result = "";

    //             if($data->department == 1){
    //                 $result .= 'TS';
    //             }
    //             else if($data->department == 2){
    //                 $result .= 'PPS';
    //             }
    //             else if($data->department == 3){
    //                 $result .= 'CN';
    //             }
    //             else{
    //                 $result .= 'YF';
    //             }

    //             return $result;
    //         })
    //         ->rawColumns(['raw_status', 'raw_action', 'raw_department', 'raw_audit_status'])
    //         ->make(true);
    // }

    /**
     * Commented as backup
     * ? JD
     */
    // public function view_audit_results(Request $request){
    //     session_start();
    //     $audit_results;
    //     $userId = $request->userId;
    //     // if(isset($request->status)){
    //         $audit_results = AuditResult::with([
    //                 'checked_by_info',
    //                 'product_line_details.product_line_info',
    //                 'section_details.section_info',
    //                 'auditor_details.user_info',
    //                 'auditee_details.user_info',
    //             ])
    //             ->where('layer', $request->layer)

    //             ->get();
    //     // }
    //     // else {
    //     //     $audit_results = AuditResult::where('status', $request->status)
    //     //     ->where('layer', $request->layer)
    //     //     ->get();
    //     // }
    //         // return $request->userId;
    //     return DataTables::of($audit_results)
    //         ->addColumn('raw_status', function($data){
    //             $result = "";

    //             if($data->status == 1){
    //                 $result .= '<span class="tag tag-success">Active</span>';
    //             }
    //             else{
    //                 $result .= '<span class="tag tag-danger">Inactive</span>';
    //             }

    //             return $result;
    //         })
    //         ->addColumn('raw_audit_status', function($data){
    //             $result = "";

    //             if($data->audit_status == 1){
    //                 $result .= '<span class="tag tag-warning">For Checking</span>';
    //             }
    //             else if($data->audit_status == 2){
    //                 $result .= '<span class="tag tag-primary">For CAPA</span>';
    //             }
    //             else if($data->audit_status == 3){
    //                 $result .= '<span class="tag tag-info">For Close Out</span>';
    //             }
    //             else if($data->audit_status == 4){
    //                 $result .= '<span class="tag tag-success">Closed</span>';
    //             }
    //             else{
    //                 $result .= '<span class="tag tag-danger">Unclosed</span>';
    //             }

    //             return $result;
    //         })
    //         ->addColumn('raw_action', function($data){
    //             $result = '';
    //             if($data->status == 1) {
    //                 $result .= '<button class="btn-actions btn btn-sm btn-success btnEditAuditResult" audit_result-id="' . $data->id . '" type="button" title="Edit" layer="' . $data->layer . '"><i class="icon-edit2"></i></button> ';

    //                 $result .= '<button class="btn-actions btn btn-sm btn-danger btnActionAuditResult" type="button" audit_result-id="' . $data->id . '" status="2" title="Archive"><i class="icon-trash2"></i></button> ';

    //                 // $result .= '<button class="btn-actions btn btn-sm btn-info btnSelectAuditResult" type="button" audit_result-id="' . $data->id . '" audit_result-name="' . $data->name . '" title="Select"><i class="icon-arrow-right"></i></button> ';
    //             }
    //             else {
    //                 $result .= '<button class="btn-actions btn btn-sm btn-warning btnActionAuditResult" type="button" audit_result-id="' . $data->id . '" status="1" title="Restore"><i class="icon-history2"></i></button> ';
    //             }

    //             return $result;
    //         })
    //         ->addColumn('raw_department', function($data){
    //             $result = "";

    //             if($data->department == 1){
    //                 $result .= 'TS';
    //             }
    //             else if($data->department == 2){
    //                 $result .= 'PPS';
    //             }
    //             else if($data->department == 3){
    //                 $result .= 'CN';
    //             }
    //             else{
    //                 $result .= 'YF';
    //             }

    //             return $result;
    //         })
    //         ->rawColumns(['raw_status', 'raw_action', 'raw_department', 'raw_audit_status'])
    //         ->make(true);
    // }

    // public function view_audit_results_layer2(Request $request){
    //     session_start();
    //     $audit_results;

    //     // if(isset($request->status)){
    //         $audit_results = AuditResult::with([
    //                 // 'checked_by_info',
    //                 'section_details.section_info',
    //                 'auditor_details.user_info',
    //                 'auditee_details.user_info',
    //             ])
    //             ->where('layer', 2)
    //             ->get();
    //     // }
    //     // else {
    //     //     $audit_results = AuditResult::where('status', $request->status)
    //     //     ->where('layer', $request->layer)
    //     //     ->get();
    //     // }

    //     return DataTables::of($audit_results)
    //         ->addColumn('raw_status', function($data){
    //             $result = "";

    //             if($data->status == 1){
    //                 $result .= '<span class="tag tag-success">Active</span>';
    //             }
    //             else{
    //                 $result .= '<span class="tag tag-danger">Inactive</span>';
    //             }

    //             return $result;
    //         })
    //         ->addColumn('raw_audit_status', function($data){
    //             $result = "";

    //             if($data->audit_status == 1){
    //                 $result .= '<span class="tag tag-warning">For Checking</span>';
    //             }
    //             else if($data->audit_status == 2){
    //                 $result .= '<span class="tag tag-primary">For CAPA</span>';
    //             }
    //             else if($data->audit_status == 3){
    //                 $result .= '<span class="tag tag-info">For Close Out</span>';
    //             }
    //             else if($data->audit_status == 4){
    //                 $result .= '<span class="tag tag-success">Closed</span>';
    //             }
    //             else{
    //                 $result .= '<span class="tag tag-danger">Unclosed</span>';
    //             }

    //             return $result;
    //         })
    //         ->addColumn('raw_action', function($data){
    //             $result = '';
    //             if($data->status == 1) {
    //                 $result .= '<button class="btn-actions btn btn-sm btn-success btnEditAuditResult" audit_result-id="' . $data->id . '" type="button" title="Edit" layer="' . $data->layer . '"><i class="icon-edit2"></i></button> ';

    //                 $result .= '<button class="btn-actions btn btn-sm btn-danger btnActionAuditResult" type="button" audit_result-id="' . $data->id . '" status="2" title="Archive"><i class="icon-trash2"></i></button> ';

    //                 // $result .= '<button class="btn-actions btn btn-sm btn-info btnSelectAuditResult" type="button" audit_result-id="' . $data->id . '" audit_result-name="' . $data->name . '" title="Select"><i class="icon-arrow-right"></i></button> ';
    //             }
    //             else {
    //                 $result .= '<button class="btn-actions btn btn-sm btn-warning btnActionAuditResult" type="button" audit_result-id="' . $data->id . '" status="1" title="Restore"><i class="icon-history2"></i></button> ';
    //             }

    //             return $result;
    //         })
    //         ->addColumn('raw_department', function($data){
    //             $result = "";

    //             if($data->department == 1){
    //                 $result .= 'TS';
    //             }
    //             else if($data->department == 2){
    //                 $result .= 'PPS';
    //             }
    //             else if($data->department == 3){
    //                 $result .= 'CN';
    //             }
    //             else{
    //                 $result .= 'YF';
    //             }

    //             return $result;
    //         })
    //         ->rawColumns(['raw_status', 'raw_action', 'raw_department', 'raw_audit_status'])
    //         ->make(true);
    // }

    public function save_audit_result(Request $request){
        date_default_timezone_set('Asia/Manila');
        session_start();

        $data = $request->all();

        // $rules = [];
        // return gettype($request->audit_status);
        // if($request->audit_status == 6){
        //     echo '6';
        // }else{
        //     echo 'not 6';
        // }
        // return;


            // if audit_result_id is not set/empty then execute insert operation else update
            if(!isset($request->audit_result_id)){
                $rules = [
                    'layer' => 'required',
                    'audit_date' => 'required',
                    // 'issued_date' => 'required',
                    // 'checked_by' => 'required', // 10-16-2021 - Jannus
                    'due_date' => 'required', // 06-08-2022 - JD
                    'issued_date' => 'required',
                    'product_line_id' => 'required',
                    // 'due_date' => 'required',
                    'commendation' => 'required',
                    'auditors' => 'required',
                    'auditees' => 'required',
                    'audit_status' => 'required',
                    'section_id' => 'required',
                ];

                if($request->audit_status == 6){
                    $rules['due_date'] = '';
                }else{
                    $rules['due_date'] = 'required';
                }

                if($request->layer == 1){
                    $rules['checked_by'] = 'required';
                    $rules['due_date'] = ''; // 10-26-2021 - Jannus
                    $rules['issued_date'] = ''; // 10-26-2021 - Jannus
                }
                // else{
                //     $rules['due_date'] = 'required';
                //     $rules['issued_date'] = 'required';
                // }

                $messages = [
                    'layer.required' => 'Layer field is required.',
                    'audit_date.required' => 'Audit Date field is required.',
                    'issued_date.required' => 'Issued Date field is required.',
                    'checked_by.required' => 'Checked By field is required.',
                    'product_line_id.required' => 'Department / Section field is required.',
                    // 'due_date.required' => 'Due Date field is required.', // 06-08-2022 - JD
                    'commendation.required' => 'Commendation field is required.',
                    'auditors.required' => 'Auditors field is required.',
                    'auditees.required' => 'Auditees field is required.',
                    'audit_status.required' => 'Audit Status field is required.',
                    'section_id.required' => 'Responsible Section field is required.',
                ];

                if($request->audit_status == 6){
                    $messages['due_date.required'] = '';
                }

                $validator = Validator::make($data, $rules, $messages);

                if ($validator->passes()) {
                    DB::beginTransaction();
                    try{
                        $auditee_manual = '';

                        if(isset($request->auditee_manual)){
                            $auditee_manual = implode(',', $request->auditee_manual);
                        }

                        $audit_result_id = AuditResult::insertGetId([
                                'layer' => $request->layer,
                                'audit_date' => $request->audit_date,
                                'issued_date' => $request->issued_date,
                                'checked_by' => $request->checked_by,
                                // 'section_id' => $request->section_id,
                                'due_date' => $request->due_date,
                                // 'auditors' => $request->auditors,
                                // 'auditees' => $request->auditees,
                                'auditee_manual' => $auditee_manual,
                                // 'factor_id' => $request->factor_id,
                                'commendation' => $request->commendation,
                                'audit_status' => $request->audit_status,
                                'status' => 1,
                                'update_version' => 1,
                                'created_by' => $_SESSION["rapidx_user_id"],
                                'last_updated_by' => $_SESSION["rapidx_user_id"],
                                'created_at' => date('Y-m-d H:i:s'),
                                'updated_at' => date('Y-m-d H:i:s'),
                            ]
                        );

                        if(isset($request->product_line_id)){
                            if(count($request->product_line_id) > 0){
                                for($index = 0; $index < count($request->product_line_id); $index++){
                                    ARProductLine::insert([
                                        'audit_result_id' => $audit_result_id,
                                        'product_line_id' => $request->product_line_id[$index],
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        if(isset($request->section_id)){
                            if(count($request->section_id) > 0){
                                for($index = 0; $index < count($request->section_id); $index++){
                                    ARSection::insert([
                                        'audit_result_id' => $audit_result_id,
                                        'section_id' => $request->section_id[$index],
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        if(isset($request->auditors)){
                            if(count($request->auditors) > 0){
                                for($index = 0; $index < count($request->auditors); $index++){
                                    ARAuditor::insert([
                                        'audit_result_id' => $audit_result_id,
                                        'user_id' => $request->auditors[$index],
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        if(isset($request->auditees)){
                            if(count($request->auditees) > 0){
                                for($index = 0; $index < count($request->auditees); $index++){
                                    ARAuditee::insert([
                                        'audit_result_id' => $audit_result_id,
                                        'user_id' => $request->auditees[$index],
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        if(isset($request->email_recepient_attention)){
                            if(count($request->email_recepient_attention) > 0){
                                for($index = 0; $index < count($request->email_recepient_attention); $index++){
                                    AREmailRecipient::insert([
                                        'audit_result_id' => $audit_result_id,
                                        'user_id' => $request->email_recepient_attention[$index],
                                        'type' => 1,
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        // $arr = array();
                        if(isset($request->email_recepient_cc)){
                            if(count($request->email_recepient_cc) > 0){
                                for($index = 0; $index < count($request->email_recepient_cc); $index++){
                                    AREmailRecipient::insert([
                                        'audit_result_id' => $audit_result_id,
                                        'user_id' => $request->email_recepient_cc[$index],
                                        'type' => 2,
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);

                                    // $email_arr = array_push($arr, $request->email_recepient_cc[$index]);
                                }
                            }
                        }
                        // return print_r($email_arr);
                        // exit();

                        // $this->send_audit_result_email($audit_result_id, $request->select_layer); // Old commented by -JD 05-17-2022

                        DB::commit();
                        return response()->json(['result' => 1, 'audit_result_id' => $audit_result_id]);
                    }
                    catch(\Exeption $e){
                        DB::rollback();
                        return response()->json(['result' => 0, 'exception' => $e]);
                    }
                }
                else{
                    return response()->json(['result' => '0', 'error' => $validator->messages()]);
                }
            }
            else{
                $rules = [
                    // 'layer' => 'required',
                    'audit_date' => 'required',
                    // 'issued_date' => 'required',
                    // 'checked_by' => 'required',
                    'due_date' => 'required',
                    'issued_date' => 'required',
                    'product_line_id' => 'required',
                    'section_id' => 'required',
                    // 'due_date' => 'required',
                    'commendation' => 'required',
                    'auditors' => 'required',
                    'auditees' => 'required',
                    'audit_status' => 'required',
                ];

                if($request->audit_status == 6){
                    $rules['due_date'] = '';
                }else{
                    $rules['due_date'] = 'required';
                }

                if($request->layer == 1){
                    $rules['checked_by'] = 'required';
                    $rules['due_date'] = ''; // 10-26-2021 - Jannus
                    $rules['issued_date'] = ''; // 10-26-2021 - Jannus
                }
                // else{
                //     $rules['issued_date'] = 'required';
                //     $rules['due_date'] = 'required';
                // }

                $messages = [
                    // 'layer.required' => 'Layer field is required.',
                    'audit_date.required' => 'Audit Date field is required.',
                    'issued_date.required' => 'Issued Date field is required.',
                    'checked_by.required' => 'Checked By field is required.',
                    'product_line_id.required' => 'Department / Section field is required.',
                    'section_id.required' => 'Responsible Section field is required.',
                    // 'due_date.required' => 'Due Date field is required.',
                    'commendation.required' => 'Commendation field is required.',
                    'auditors.required' => 'Auditors field is required.',
                    'auditees.required' => 'Auditees field is required.',
                    'audit_status.required' => 'Audit Status field is required.',
                ];

                if($request->audit_status == 6){
                    $messages['due_date.required'] = '';
                }

                $validator = Validator::make($data, $rules, $messages);

                if ($validator->passes()) {
                    DB::beginTransaction();
                    try{
                        $auditee_manual = '';

                        if(isset($request->auditee_manual)){
                            $auditee_manual = implode(',', $request->auditee_manual);
                        }

                        AuditResult::where('id', $request->audit_result_id)
                            ->increment('update_version', 1,
                            [
                                // 'layer' => $request->layer,
                                'audit_date' => $request->audit_date,
                                'issued_date' => $request->issued_date,
                                'checked_by' => $request->checked_by,
                                'due_date' => $request->due_date,
                                'auditee_manual' => $auditee_manual,
                                'commendation' => $request->commendation,
                                'audit_status' => $request->audit_status,
                                'last_updated_by' => $_SESSION["rapidx_user_id"],
                                'updated_at' => date('Y-m-d H:i:s'),
                            ]);

                        ARProductLine::where('audit_result_id', $request->audit_result_id)
                        ->delete();
                        if(isset($request->product_line_id)){
                            if(count($request->product_line_id) > 0){
                                for($index = 0; $index < count($request->product_line_id); $index++){
                                    ARProductLine::insert([
                                        'audit_result_id' => $request->audit_result_id,
                                        'product_line_id' => $request->product_line_id[$index],
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        ARSection::where('audit_result_id', $request->audit_result_id)
                        ->delete();
                        if(isset($request->section_id)){
                            if(count($request->section_id) > 0){
                                for($index = 0; $index < count($request->section_id); $index++){
                                    ARSection::insert([
                                        'audit_result_id' => $request->audit_result_id,
                                        'section_id' => $request->section_id[$index],
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        ARAuditor::where('audit_result_id', $request->audit_result_id)
                        ->delete();
                        if(isset($request->auditors)){
                            if(count($request->auditors) > 0){
                                for($index = 0; $index < count($request->auditors); $index++){
                                    ARAuditor::insert([
                                        'audit_result_id' => $request->audit_result_id,
                                        'user_id' => $request->auditors[$index],
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        ARAuditee::where('audit_result_id', $request->audit_result_id)
                        ->delete();
                        if(isset($request->auditees)){
                            if(count($request->auditees) > 0){

                                for($index = 0; $index < count($request->auditees); $index++){
                                    ARAuditee::insert([
                                        'audit_result_id' => $request->audit_result_id,
                                        'user_id' => $request->auditees[$index],
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        AREmailRecipient::where('audit_result_id', $request->audit_result_id)
                        ->delete();

                        if(isset($request->email_recepient_attention)){
                            if(count($request->email_recepient_attention) > 0){
                                for($index = 0; $index < count($request->email_recepient_attention); $index++){
                                    AREmailRecipient::insert([
                                        'audit_result_id' => $request->audit_result_id,
                                        'user_id' => $request->email_recepient_attention[$index],
                                        'type' => 1,
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        if(isset($request->email_recepient_cc)){
                            if(count($request->email_recepient_cc) > 0){
                                for($index = 0; $index < count($request->email_recepient_cc); $index++){
                                    AREmailRecipient::insert([
                                        'audit_result_id' => $request->audit_result_id,
                                        'user_id' => $request->email_recepient_cc[$index],
                                        'type' => 2,
                                        'status' => 1,
                                        'created_by' => $_SESSION["rapidx_user_id"],
                                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                                        'created_at' => date('Y-m-d H:i:s'),
                                        'updated_at' => date('Y-m-d H:i:s'),
                                    ]);
                                }
                            }
                        }

                        // $this->send_audit_result_email($request->audit_result_id);

                        DB::commit();
                        return response()->json(['result' => 1, 'audit_result_id' => $request->audit_result_id]);
                    }
                    catch(\Exeption $e){
                        DB::rollback();
                        return response()->json(['result' => 0, 'exception' => $e]);
                    }
                }
                else{
                    return response()->json(['result' => '0', 'error' => $validator->messages()]);
                }
            }
        // }

        // else if($request->layer == 2){ // LAYER 2

        // }

        // else if($request->layer == 3){ // LAYER 3

        // }
    }

    public function get_audit_result_by_id(Request $request){
        session_start();
        $data = AuditResult::where('id', $request->audit_result_id)
                ->with([
                    'checked_by_info',
                    'product_line_details.product_line_info',
                    'section_details.section_info',
                    'auditor_details.user_info',
                    'auditee_details.user_info',
                    'ar_email_recipient_details.user_info',
                ])
                ->first();

        return response()->json(['data' => $data]);
    }

    public function action_audit_result(Request $request){
        date_default_timezone_set('Asia/Manila');
        session_start();

        $data = $request->all();

        $rules = [
            'audit_result_id' => ['required'],
            'status' => ['required'],
        ];

        $validator = Validator::make($data, $rules);

        if ($validator->passes()) {
            DB::beginTransaction();
            try{

                AuditResult::where('id', $request->audit_result_id)
                    ->update([
                        'status' => $request->status,
                        'last_updated_by' => $_SESSION["rapidx_user_id"],
                        'updated_at' => date('Y-m-d H:i:s'),
                    ]
                );

                DB::commit();
                return response()->json(['result' => 1]);
            }
            catch(\Exeption $e){
                DB::rollback();
                return response()->json(['result' => 0]);
            }
        }
        else{
            return response()->json(['result' => '0', 'error' => $validator->messages()]);
        }
    }

    public function get_cbo_audit_results_by_stat(Request $request){
        date_default_timezone_set('Asia/Manila');

        $search = $request->search;

        if($search == ''){
            $datas = [];
        }
        else{
            $datas = AuditResult::orderby('name','asc')->select('id','name')
                        ->where('name', 'like', '%' . $search . '%')
                        ->where('status', 1)
                        ->get();
        }

        $response = array();
        $response[] = array(
            "id" => '',
            "text" => '',
        );

        foreach($datas as $data){
            $response[] = array(
                "id" => $data->id,
                "text" => $data->name,
            );
        }

        echo json_encode($response);
        exit;
    }

    public function export_audit_result(Request $request){
        session_start();
        $data = AuditResult::where('id', $request->audit_result_id)
                ->with([
                    'checked_by_info',
                    'product_line_details.product_line_info',
                    'auditor_details.user_info',
                    'auditee_details.user_info',
                    'ar_email_recipient_details.user_info', // rapidx users
                    'ar_finding_details',
                    'ar_finding_details.in_charge_details.user_info', // rapidx user
                    'ar_finding_details.classification_info',
                ])
                ->first();

        // MAN
        $factor_item_list = FactorItemList::where('factor_id', 1)->with([
            'factor'
        ])->orderBy('name', 'ASC')->get();

        // Machine
        $factor_item_list_machine = FactorItemList::where('factor_id', 3)->with([
            'factor'
        ])->orderBy('name', 'ASC')->get();

        // Method
        $factor_item_list_method = FactorItemList::where('factor_id', 2)->with([
            'factor'
        ])->orderBy('name', 'ASC')->get();

        // Material
        $factor_item_list_material = FactorItemList::where('factor_id', 5)->with([
            'factor'
        ])->orderBy('name', 'ASC')->get();

        // Flow of non-conforming
        $factor_item_list_non_forming = FactorItemList::where('factor_id', 4)->with([
            'factor'
        ])->orderBy('name', 'ASC')->get();
        // dd($factor_item_list);

        return Excel::download(new AuditResultsExport($data, $factor_item_list, $factor_item_list_machine, $factor_item_list_method, $factor_item_list_material, $factor_item_list_non_forming), 'QC Patrol Audit Result - ' . date('Y-m-d H:i:s') . '.xlsx');
        return view('exports.audit_results')->with(['data' => $data]);
    }

    // public function send_audit_result_email(Request $request){
    public function send_audit_result_email($audit_result_id, $select_layer){
        // session_start();
        $audit_result = AuditResult::where('id', $audit_result_id)
                ->with([
                    'checked_by_info', // rapidx users
                    'product_line_details.product_line_info',
                    'section_details.section_info',
                    'auditor_details.user_info',
                    'auditee_details.user_info',
                    'ar_email_recipient_details.user_info', // rapidx users
                    'ar_finding_details',
                    'ar_finding_details.in_charge_details.user_info',
                    'ar_finding_details.classification_info',
                    'ar_finding_details.factor_info',
                    'ar_finding_details.factor_item_list_info',
                ])
                ->first();


        /*
            Commented by -JD
        */
        // return $audit_result->checked_by;
        // exit();

        /*
            Commented by -JD
        */
        // $to_emailarr = [];
        // $to_emailarr = collect($audit_result->ar_email_recipient_details)->where('audit_result_id', $audit_result_id)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();

        // print_r($to_emailarr);
        // exit();

        $to_email = [];
        $cc_email = [];

        // return $data;

        $subject = '';
        if($audit_result != null){
            if($audit_result->audit_status == 1){
                $subject = 'For Checking';
                if($select_layer== 1){ // Added by -JD
                    if($audit_result->checked_by_info->email != null){ // rapidx users email
                        $to_email[] = $audit_result->checked_by_info->email;
                        // $cc_email = collect($audit_result->ar_email_recipient_details)->where('type', 2)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();
                    }
                }
                /*
                    Adding $to_email & $cc_email for Layer 2 & 3
                    Added by -JD
                */
                else{
                    $to_email = collect($audit_result->ar_email_recipient_details)->where('type', 1)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();
                    $cc_email = collect($audit_result->ar_email_recipient_details)->where('type', 2)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();
                }

            }
            else{
                if($audit_result->audit_status == 2){
                    $subject = 'For CAPA';
                }
                else if($audit_result->audit_status == 3){
                    $subject = 'For Close Out';
                }
                else if($audit_result->audit_status == 4){
                    // $subject = 'Closed</span';
                    $subject = 'Closed';
                }
                else if($audit_result->audit_status == 5){
                    // $subject = 'Unclosed</span';
                    $subject = 'Unclosed';
                }
                else{
                    $subject = 'No Findings';
                }

                // user_info() in AREmailRecipient Model - rapidx users
                $to_email = collect($audit_result->ar_email_recipient_details)->where('type', 1)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();
                $cc_email = collect($audit_result->ar_email_recipient_details)->where('type', 2)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();
            }
        }

        // $subject = 'For Checking';

        // print_r($to_email);
        // echo $to_email;
        // exit();
        // return ;

        $data = [
            'data' => $audit_result,
            'subject' => $subject,
        ];

        // $email_recipients = [
        //     'to' => $to_email,
        //     'cc' => $cc_email,
        // ];

        $email_recipients = [
            'to' => $to_email,
            'cc' => $cc_email,
        ];

        // return $to_email;

        // return view('mail.audit_result.for_capa')->with(['data' => $audit_result, 'subject' => $subject]);

        try {

            $mail = Mail::send('mail.audit_result.for_capa', $data, function($message) use ($email_recipients, $data) {
                $message->to($email_recipients['to'])
                ->cc($email_recipients['cc'])
                ->bcc('jdomingo@pricon.ph')
                ->subject('QC Patrols - ' . $data['subject']);
            });

            return response()->json(['result' => 1]);
        }
        catch (Exception $e) {
            return response()->json(['result' => 0]);
        }

    }


    public function btn_send_audit_result_email(Request $request){
        // session_start();
        $audit_result = AuditResult::where('id', $request->audit_result_id)
                ->with([
                    'checked_by_info',
                    'product_line_details.product_line_info',
                    'section_details.section_info',
                    'auditor_details.user_info',
                    'auditee_details.user_info',
                    'ar_email_recipient_details.user_info',
                    'ar_finding_details',
                    'ar_finding_details.in_charge_details.user_info',
                    'ar_finding_details.classification_info',
                    'ar_finding_details.factor_info',
                    'ar_finding_details.factor_item_list_info',
                ])
                ->first();

        // return $audit_result;

        $to_email = [];
        $cc_email = [];

        // $to_email = [
        //     'apbautista@pricon.ph',
        //     'dfamandoron@pricon.ph',
        //     // 'jgsulit@pricon.ph',
        //     // 'jnpineda@pricon.ph',
        // ];

        // return $data;

        $subject = '';

        if($audit_result != null){
            // if "For Checking" then send to users email only
            if($audit_result->audit_status == 1){
                $subject = 'For Checking';
                if($request->layer == 1){
                    if($audit_result->checked_by_info->email != null){ // rapidx users email
                        $to_email[] = $audit_result->checked_by_info->email;
                        // $cc_email = collect($audit_result->ar_email_recipient_details)->where('type', 2)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();
                    }
                }
                /*
                    Adding $to_email & $cc_email for Layer 2 & 3
                    Added by -JD
                */
                else{
                    $to_email = collect($audit_result->ar_email_recipient_details)->where('type', 1)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();
                    $cc_email = collect($audit_result->ar_email_recipient_details)->where('type', 2)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();
                }

            }
            // else then send to email recipients
            else{
                if($audit_result->audit_status == 2){
                    $subject = 'For CAPA';
                }
                else if($audit_result->audit_status == 3){
                    $subject = 'For Close Out';
                }
                else if($audit_result->audit_status == 4){
                    $subject = 'Closed';
                }
                else if($audit_result->audit_status == 5){
                    // $subject = 'Unclosed</span';
                    $subject = 'Unclosed';
                }
                else{
                    $subject = 'No Findings';
                }

                // user_info() in AREmailRecipient Model - rapidx users
                $to_email = collect($audit_result->ar_email_recipient_details)->where('type', 1)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();
                $cc_email = collect($audit_result->ar_email_recipient_details)->where('type', 2)->pluck('user_info')->flatten(1)->pluck('email')->flatten(1)->toArray();
            }
        }

        // $subject = 'For Checking';

        // echo($subject);
        // print_r($to_email);
        // print_r($cc_email);
        // exit();
        // return ;

        $data = [
            'data' => $audit_result,
            'subject' => $subject,
        ];

        // $email_recipients = [
        //     'to' => $to_email,
        //     'cc' => $cc_email,
        // ];

        $email_recipients = [
            'to' => $to_email,
            'cc' => $cc_email,
        ];

        // print_r($to_email);
        // print_r($cc_email);
        // exit();

        // return $to_email;

        // return view('mail.audit_result.for_capa')->with(['data' => $audit_result, 'subject' => $subject]);

        try {

            $mail = Mail::send('mail.audit_result.for_capa', $data, function($message) use ($email_recipients, $data) {
                $message->to($email_recipients['to'])
                ->cc($email_recipients['cc'])
                ->subject('QC Patrol - ' . $data['subject']);
            });

            return response()->json(['result' => 1]);
        }
        catch (Exception $e) {
            return response()->json(['result' => 0]);
        }

    }
}
