import { useEffect, useMemo, useState } from "react";
import { supabase } from "./supabaseClient";

type Profile={id:string;username:string;full_name:string;is_admin:boolean;is_active:boolean;auth_email?:string|null;can_view_surveys?:boolean;can_view_answers?:boolean;created_at?:string};
type Survey={id:number;title:string;description:string|null;is_active:boolean;created_at:string;updated_at?:string};
type Question={id:number;survey_id:number;question_text:string;question_order:number};
type Option={id:number;question_id:number;option_text:string;option_order:number};
type Vote={id:number;survey_id:number;question_id:number;option_id:number;user_id:string;created_at:string};
type Suggestion={id:number;user_id:string;survey_id:number;suggestion_text:string;created_at:string;profiles?:{username:string;full_name:string}};

export default function App(){
 const [profile,setProfile]=useState<Profile|null>(null),[loading,setLoading]=useState(true),[loginLoading,setLoginLoading]=useState(false),[username,setUsername]=useState(""),[password,setPassword]=useState(""),[message,setMessage]=useState("");
 const [menu,setMenu]=useState("dashboard"),[surveys,setSurveys]=useState<Survey[]>([]),[questions,setQuestions]=useState<Question[]>([]),[options,setOptions]=useState<Option[]>([]),[users,setUsers]=useState<Profile[]>([]),[votes,setVotes]=useState<Vote[]>([]),[suggestions,setSuggestions]=useState<Suggestion[]>([]),[selected,setSelected]=useState<Survey|null>(null);
 const [modal,setModal]=useState<"survey"|"question"|"option"|"user"|"answer"|null>(null),[busy,setBusy]=useState(false),[notice,setNotice]=useState("");
 const [surveyForm,setSurveyForm]=useState({id:0,title:"",description:"",is_active:true});
 const [qForm,setQForm]=useState({id:0,text:""}),[oForm,setOForm]=useState({id:0,question_id:0,text:""});
 const [sharedOptions,setSharedOptions]=useState<string[]>([
  "می‌پسندم",
  "نمی‌پسندم",
  "نظری ندارم"
]);
 const [uForm,setUForm]=useState({id:"",username:"",full_name:"",email:"",password:"",is_admin:false,is_active:true,can_view_surveys:true,can_view_answers:true});
 const [answerSurvey,setAnswerSurvey]=useState<Survey|null>(null),[answers,setAnswers]=useState<Record<number,number>>({}),[suggestion,setSuggestion]=useState("");
 const [completions,setCompletions]=useState<any[]>([]),[viewUser,setViewUser]=useState<Profile|null>(null),[viewSurvey,setViewSurvey]=useState<Survey|null>(null);
 const [userSearch,setUserSearch]=useState(""),[answerSurveyFilter,setAnswerSurveyFilter]=useState(0);

 const profileOf=async(id:string)=>{const {data,error}=await supabase.from("profiles").select("id,username,full_name,is_admin,is_active,auth_email,can_view_surveys,can_view_answers,created_at").eq("id",id).maybeSingle();if(error){setMessage("خطای پروفایل: "+error.message);return null}if(!data){setMessage("اطلاعات کاربر پیدا نشد.");return null}if(!data.is_active){await supabase.auth.signOut();setMessage("حساب کاربری شما غیرفعال است.");return null}setProfile(data);return data};
 useEffect(()=>{let alive=true;(async()=>{const {data:{session}}=await supabase.auth.getSession();if(alive&&session?.user)await profileOf(session.user.id);if(alive)setLoading(false)})();const {data}=supabase.auth.onAuthStateChange(async(_,s)=>{if(s?.user)await profileOf(s.user.id);else setProfile(null);setLoading(false)});return()=>{alive=false;data.subscription.unsubscribe()}},[]);
 const login=async(e:React.FormEvent)=>{e.preventDefault();setMessage("");if(!username.trim()||!password)return setMessage("نام کاربری و رمز عبور را وارد کنید.");setLoginLoading(true);try{const {data:email,error}=await supabase.rpc("get_login_email",{p_username:username.trim()});if(error||!email)return setMessage("نام کاربری یا رمز عبور اشتباه است.");const {data,error:le}=await supabase.auth.signInWithPassword({email,password});if(le||!data.user)return setMessage("نام کاربری یا رمز عبور اشتباه است.");await profileOf(data.user.id)}finally{setLoginLoading(false)}};
 const logout=async()=>{await supabase.auth.signOut();setProfile(null);setUsername("");setPassword("");setMenu("dashboard")};
 const loadSurveys=async()=>{const {data,error}=await supabase.from("surveys").select("id,title,description,is_active,created_at,updated_at").order("created_at",{ascending:false});if(error)setNotice(error.message);else setSurveys(data||[])};
 const loadUsers=async()=>{const {data,error}=await supabase.from("profiles").select("id,username,full_name,is_admin,is_active,auth_email,can_view_surveys,can_view_answers,created_at").order("created_at",{ascending:false});if(error)setNotice(error.message);else setUsers(data||[])};
 const loadVotes=async()=>{const {data}=await supabase.from("survey_votes").select("id,survey_id,question_id,option_id,user_id,created_at");setVotes(data||[])};
 const loadCompletions=async()=>{const {data,error}=await supabase.from("survey_completions").select("id,user_id,survey_id,created_at").order("created_at",{ascending:false});if(!error)setCompletions(data||[])};
 const loadSuggestions = async () => {
  const { data, error } = await supabase
    .from("suggestions")
    .select("id,user_id,survey_id,suggestion_text,created_at,profiles(username,full_name)")
    .order("created_at", { ascending: false });

  if (error) {
    setNotice(error.message);
    return;
  }

  const normalizedSuggestions: Suggestion[] = (data || []).map((item: any) => ({
    id: item.id,
    user_id: item.user_id,
    survey_id: item.survey_id,
    suggestion_text: item.suggestion_text,
    created_at: item.created_at,
    profiles: Array.isArray(item.profiles)
      ? (item.profiles[0] || { username: "", full_name: "" })
      : (item.profiles || { username: "", full_name: "" }),
  }));

  setSuggestions(normalizedSuggestions);
};
 const structure=async(sid:number)=>{const {data:q,error:qe}=await supabase.from("survey_questions").select("id,survey_id,question_text,question_order").eq("survey_id",sid).order("question_order");if(qe){setNotice(qe.message);return}setQuestions(q||[]);const ids=(q||[]).map(x=>x.id);if(!ids.length)return setOptions([]);const {data:o}=await supabase.from("survey_options").select("id,question_id,option_text,option_order").in("question_id",ids).order("option_order");setOptions(o||[])};
 useEffect(()=>{if(!profile)return;loadSurveys();loadCompletions();if(profile.is_admin){loadUsers();loadVotes();loadSuggestions()}},[profile]);
 const saveSurvey=async(e:React.FormEvent)=>{e.preventDefault();if(!surveyForm.title.trim())return;const payload={title:surveyForm.title.trim(),description:surveyForm.description.trim()||null,is_active:surveyForm.is_active,updated_at:new Date().toISOString()};const r=surveyForm.id?await supabase.from("surveys").update(payload).eq("id",surveyForm.id):await supabase.from("surveys").insert(payload);if(r.error)return setNotice(r.error.message);setModal(null);await loadSurveys()};
 const removeSurvey=async(id:number)=>{if(!confirm("نظرسنجی حذف شود؟"))return;const r=await supabase.from("surveys").delete().eq("id",id);if(r.error)setNotice(r.error.message);else await loadSurveys()};
 //const saveQuestion=async(e:React.FormEvent)=>{e.preventDefault();if(!selected||!qForm.text.trim())return;const order=questions.reduce((m,x)=>Math.max(m,x.question_order),0)+1;const r=qForm.id?await supabase.from("survey_questions").update({question_text:qForm.text.trim()}).eq("id",qForm.id):await supabase.from("survey_questions").insert({survey_id:selected.id,question_text:qForm.text.trim(),question_order:order});if(r.error){setNotice("خطا در ذخیره سؤال: "+r.error.message);return}setModal(null);await structure(selected.id)};
 const saveQuestion=async(e:React.FormEvent)=>{
  e.preventDefault();

  if(!selected||!qForm.text.trim())return;

  if(qForm.id){
    const r=await supabase
      .from("survey_questions")
      .update({
        question_text:qForm.text.trim()
      })
      .eq("id",qForm.id);

    if(r.error){
      setNotice("خطا در ذخیره سؤال: "+r.error.message);
      return;
    }
  }else{
    const order=questions.reduce(
      (m,x)=>Math.max(m,x.question_order),
      0
    )+1;

    const r=await supabase
      .from("survey_questions")
      .insert({
        survey_id:selected.id,
        question_text:qForm.text.trim(),
        question_order:order
      })
      .select("id")
      .single();

    if(r.error){
      setNotice("خطا در ذخیره سؤال: "+r.error.message);
      return;
    }

    const values=sharedOptions
      .map(x=>x.trim())
      .filter(Boolean);

    if(values.length>0){
      const optionRows=values.map((text,index)=>({
        question_id:r.data.id,
        option_text:text,
        option_order:index+1
      }));

      const optionResult=await supabase
        .from("survey_options")
        .insert(optionRows);

      if(optionResult.error){
        setNotice(
          "سؤال ساخته شد اما گزینه‌های مشترک ذخیره نشد: "+
          optionResult.error.message
        );

        setModal(null);
        await structure(selected.id);
        return;
      }
    }
  }

  setModal(null);
  await structure(selected.id);
};
const applySharedOptions=async()=>{
  if(!selected)return;

  const values=sharedOptions
    .map(x=>x.trim())
    .filter(Boolean);

  if(!values.length){
    setNotice("حداقل یک گزینه مشترک وارد کنید.");
    return;
  }

  const qs=questions.filter(q=>q.survey_id===selected.id);

  if(!qs.length){
    setNotice("ابتدا حداقل یک سؤال برای این نظرسنجی ایجاد کنید.");
    return;
  }

  if(!confirm("گزینه‌های فعلی همه سؤال‌ها با این گزینه‌ها جایگزین شوند؟")){
    return;
  }

  setBusy(true);

  try{
    const questionIds=qs.map(q=>q.id);

    const del=await supabase
      .from("survey_options")
      .delete()
      .in("question_id",questionIds);

    if(del.error){
      setNotice(
        "حذف گزینه‌های قبلی انجام نشد: "+
        del.error.message
      );
      return;
    }

    const rows=qs.flatMap(q =>
      values.map((text,index)=>({
        question_id:q.id,
        option_text:text,
        option_order:index+1
      }))
    );

    const ins=await supabase
      .from("survey_options")
      .insert(rows);

    if(ins.error){
      setNotice(
        "گزینه‌های مشترک ذخیره نشد: "+
        ins.error.message
      );
      return;
    }

    await structure(selected.id);

    setNotice(
      "گزینه‌های مشترک برای تمام سؤال‌ها اعمال شد."
    );
  }finally{
    setBusy(false);
  }
};
 const removeQuestion=async(id:number)=>{if(!confirm("سؤال حذف شود؟"))return;const r=await supabase.from("survey_questions").delete().eq("id",id);if(r.error)setNotice(r.error.message);else if(selected)await structure(selected.id)};
 const saveOption=async(e:React.FormEvent)=>{e.preventDefault();if(!oForm.question_id||!oForm.text.trim())return;const order=options.filter(x=>x.question_id===oForm.question_id).reduce((m,x)=>Math.max(m,x.option_order),0)+1;const r=oForm.id?await supabase.from("survey_options").update({option_text:oForm.text.trim()}).eq("id",oForm.id):await supabase.from("survey_options").insert({question_id:oForm.question_id,option_text:oForm.text.trim(),option_order:order});if(r.error){setNotice("خطا در ذخیره گزینه: "+r.error.message);return}setModal(null);if(selected)await structure(selected.id)};
 const removeOption=async(id:number)=>{if(!confirm("گزینه حذف شود؟"))return;const r=await supabase.from("survey_options").delete().eq("id",id);if(r.error)setNotice(r.error.message);else if(selected)await structure(selected.id)};
 const saveUser=async(e:React.FormEvent)=>{e.preventDefault();setBusy(true);try{if(uForm.id){const {error}=await supabase.from("profiles").update({username:uForm.username.trim(),full_name:uForm.full_name.trim(),is_admin:uForm.is_admin,is_active:uForm.is_active,can_view_surveys:uForm.can_view_surveys,can_view_answers:uForm.can_view_answers}).eq("id",uForm.id);if(error)return setNotice(error.message)}else{if(!uForm.username||!uForm.email||!uForm.password)return setNotice("نام کاربری، ایمیل و رمز عبور الزامی است.");const {data,error}=await supabase.functions.invoke("create-user",{body:{username:uForm.username.trim(),full_name:uForm.full_name.trim(),email:uForm.email.trim(),password:uForm.password,is_admin:uForm.is_admin,is_active:uForm.is_active,can_view_surveys:uForm.can_view_surveys,can_view_answers:uForm.can_view_answers}});if(error)return setNotice("ایجاد حساب انجام نشد: "+error.message);if(data?.error)return setNotice("ایجاد حساب انجام نشد: "+data.error)}setModal(null);await loadUsers()}finally{setBusy(false)}};
 const toggleUser=async(u:Profile,field:string)=>{const {error}=await supabase.from("profiles").update({[field]:!(u as any)[field]}).eq("id",u.id);if(error)setNotice(error.message);else loadUsers()};
 const answer=async()=>{if(!profile||!answerSurvey)return;const already= completions.some((c:any)=>c.survey_id===answerSurvey.id&&c.user_id===profile.id);if(already){setNotice("شما قبلاً در این نظرسنجی شرکت کرده‌اید.");return}const qs=questions.filter(q=>q.survey_id===answerSurvey.id);if(qs.some(q=>!answers[q.id]))return setNotice("به همه سؤال‌ها پاسخ دهید.");setBusy(true);const {data:done}=await supabase.from("survey_completions").select("id").eq("survey_id",answerSurvey.id).eq("user_id",profile.id).maybeSingle();if(done){setNotice("شما قبلاً در این نظرسنجی شرکت کرده‌اید.");setBusy(false);return}const {error}=await supabase.from("survey_votes").insert(qs.map(q=>({survey_id:answerSurvey.id,question_id:q.id,option_id:answers[q.id],user_id:profile.id})));if(error){setNotice(error.message);setBusy(false);return}const c=await supabase.from("survey_completions").insert({survey_id:answerSurvey.id,user_id:profile.id});if(c.error){setNotice(c.error.message);setBusy(false);return}if(suggestion.trim())await supabase.from("suggestions").insert({user_id:profile.id,survey_id:answerSurvey.id,suggestion_text:suggestion.trim()});await loadCompletions();setAnswerSurvey(null);setSuggestion("");setAnswers({});setBusy(false);setNotice("پاسخ با موفقیت ثبت شد.")};
 const stats=useMemo(()=>{if(!selected)return[];return questions.filter(q=>q.survey_id===selected.id).map(q=>{const os=options.filter(o=>o.question_id===q.id).map(o=>{const count=votes.filter(v=>v.survey_id===selected.id&&v.question_id===q.id&&v.option_id===o.id).length;return {...o,count}});const total=os.reduce((a,x)=>a+x.count,0);return{q,os:os.map(o=>({...o,pct:total?Math.round(o.count*100/total):0})),total}})},[selected,questions,options,votes]);
 if(loading)return <div className="app"><div className="loading-screen">در حال بارگذاری...</div></div>;
 if(!profile)return <div className="app" dir="rtl"><div className="login-card"><div className="logo">✓</div><h1>ورود به سامانه</h1><p className="subtitle">برای ادامه وارد حساب کاربری خود شوید</p><form onSubmit={login}><label>نام کاربری</label><input value={username} onChange={e=>setUsername(e.target.value)} /><label>رمز عبور</label><input type="password" value={password} onChange={e=>setPassword(e.target.value)}/><button disabled={loginLoading}>{loginLoading?"در حال ورود...":"ورود"}</button></form>{message&&<div className="error-message">{message}</div>}</div></div>;
 if(!profile.is_admin)return <UserView profile={profile} surveys={surveys.filter(s=>s.is_active)} questions={questions} options={options} answerSurvey={answerSurvey} setAnswerSurvey={async (s: Survey) => {
  setAnswerSurvey(s);
  setAnswers({});
  setSuggestion("");
  await structure(s.id);
}} answers={answers} setAnswers={setAnswers} suggestion={suggestion} setSuggestion={setSuggestion} answer={answer} busy={busy} logout={logout} notice={notice} completions={completions} />;
 return <AdminView {...{profile,menu,setMenu,surveys,users,votes,suggestions,selected,setSelected,questions,options,stats,notice,setNotice,modal,setModal,surveyForm,setSurveyForm,qForm,setQForm,sharedOptions,setSharedOptions,oForm,setOForm,uForm,setUForm,saveSurvey,removeSurvey,saveQuestion,applySharedOptions,removeQuestion,saveOption,removeOption,saveUser,toggleUser,loadSurveys,loadVotes,loadSuggestions,loadCompletions,structure,answerSurvey,setAnswerSurvey,answers,setAnswers,suggestion,setSuggestion,busy,setBusy,logout,completions,viewUser,setViewUser,viewSurvey,setViewSurvey,userSearch,setUserSearch,answerSurveyFilter,setAnswerSurveyFilter}}/>;
}

function UserView(p:any){return <div className="user-page" dir="rtl"><header className="user-header"><div><strong>سامانه نظرسنجی</strong><span>پنل کاربری</span></div><button type="button" onClick={() => p.logout()} style={{display:"block",visibility:"visible",opacity:1,cursor:"pointer",border:"none",padding:"10px 18px",borderRadius:"10px",background:"#ef4444",color:"#fff",fontSize:"14px",fontFamily:"inherit"}}>🚪 خروج از حساب</button></header><main className="user-content"><section className="user-welcome"><div><span>خوش آمدید</span><h1>{p.profile.full_name} 👋</h1><p>نظرسنجی‌های فعال را مشاهده و در آن‌ها شرکت کنید.</p></div><div className="big-icon">🗳️</div></section>{p.profile.can_view_surveys===false?<div className="empty-state"><div>🔒</div><h3>دسترسی به نظرسنجی‌ها غیرفعال است</h3></div>:<div className="survey-user-grid">{!p.surveys.length?<div className="empty-state"><div>📋</div><h3>نظرسنجی فعالی وجود ندارد</h3></div>:p.surveys.map((s:Survey)=><article className="user-survey-card" key={s.id}><div className="card-icon">📋</div><h3>{s.title}</h3><p>{s.description||"برای شرکت کلیک کنید."}</p>{p.completions.some((c:any)=>c.survey_id===s.id&&c.user_id===p.profile.id)?<><span className="status active">✓ قبلاً شرکت کرده‌اید</span><button disabled>پاسخ ثبت شده</button></>:<button onClick={()=>p.setAnswerSurvey(s)}>شرکت در نظرسنجی</button>}</article>)}</div>}{p.answerSurvey&&<div className="modal-backdrop"><div className="modal large-modal"><div className="modal-header"><div><h2>{p.answerSurvey.title}</h2><p>{p.answerSurvey.description}</p></div><button className="close-button" onClick={()=>p.setAnswerSurvey(null)}>×</button></div>{p.questions.filter((q:Question)=>q.survey_id===p.answerSurvey.id).map((q:Question,i:number)=><div className="answer-question" key={q.id}><h3>{i+1}. {q.question_text}</h3>{p.options.filter((o:Option)=>o.question_id===q.id).map((o:Option)=><label className="radio-option" key={o.id}><input type="radio" name={'q'+q.id} checked={p.answers[q.id]===o.id} onChange={()=>p.setAnswers((x:any)=>({...x,[q.id]:o.id}))}/>{o.option_text}</label>)}</div>)}<div className="suggestion-box"><label>پیشنهاد کلی شما <span>(اختیاری)</span></label><textarea rows={4} value={p.suggestion} onChange={(e:any)=>p.setSuggestion(e.target.value)} placeholder="پیشنهاد شما درباره این نظرسنجی..."/></div>{p.notice&&<div className="message-box">{p.notice}</div>}<div className="form-actions"><button className="secondary-button" onClick={()=>p.setAnswerSurvey(null)}>انصراف</button><button className="primary-button" disabled={p.busy} onClick={p.answer}>{p.busy?"در حال ثبت...":"ثبت پاسخ‌ها"}</button></div></div></div>}</main></div>}

function AdminView(p:any){const nav=[['dashboard','🏠','داشبورد'],['surveys','📋','نظرسنجی‌ها'],['users','👥','کاربران'],['answers','📝','پاسخ‌ها'],['results','📊','نتایج و نمودارها'],['suggestions','💬','پیشنهادها']];return <div className="admin-layout" dir="rtl"><aside className="sidebar"><div className="brand"><div className="brand-icon">✓</div><div><strong>سامانه نظرسنجی</strong><small>پنل مدیریت</small></div></div><nav>{nav.map(n=><button key={n[0]} className={p.menu===n[0]?'menu active':'menu'} onClick={()=>p.setMenu(n[0])}><span>{n[1]}</span><span>{n[2]}</span></button>)}</nav><div className="sidebar-bottom"><div className="admin-profile"><div className="avatar">{p.profile.full_name?.charAt(0)}</div><div><strong>{p.profile.full_name}</strong><small>مدیر سیستم</small></div></div><button className="logout" onClick={p.logout}>🚪 خروج</button></div></aside><main className="admin-main"><header className="topbar"><div><h1>{nav.find(n=>n[0]===p.menu)?.[2]}</h1><p>خوش آمدید، {p.profile.full_name}</p></div><div className="topbar-avatar">{p.profile.full_name?.charAt(0)}</div></header>{p.notice&&<div className="message-box">{p.notice}</div>}{p.menu==='dashboard'&&<Dashboard p={p}/>} {p.menu==='surveys'&&<Surveys p={p}/>} {p.menu==='users'&&<Users p={p}/>} {p.menu==='answers'&&<Answers p={p}/>} {p.menu==='results'&&<Results p={p}/>} {p.menu==='suggestions'&&<Suggestions p={p}/>}<Modals p={p}/></main></div>}

function Dashboard({p}:any){return <><section className="welcome-card"><div><span className="welcome-label">مدیریت سامانه</span><h2>به پنل مدیریت خوش آمدید 👋</h2><p>نظرسنجی‌ها، کاربران، پاسخ‌ها و پیشنهادها را مدیریت کنید.</p></div><div className="welcome-icon">📊</div></section><section className="stats-grid">{[["📋","نظرسنجی‌ها",p.surveys.length],["👥","کاربران",p.users.length],["🗳️","پاسخ‌ها",p.votes.length],["💬","پیشنهادها",p.suggestions.length]].map(x=><div className="stat-card" key={x[1]}><div className="stat-icon">{x[0]}</div><div><span>{x[1]}</span><strong>{x[2]}</strong></div></div>)}</section></>}

function Surveys({p}:any){return <section className="content-card"><div className="section-header"><div><h2>مدیریت نظرسنجی‌ها</h2><p>ایجاد، حذف، سؤال و گزینه</p></div><button className="primary-button" onClick={()=>{p.setSurveyForm({id:0,title:'',description:'',is_active:true});p.setModal('survey')}}>+ نظرسنجی جدید</button></div><div className="survey-admin-list">{p.surveys.map((s:Survey)=><div className="survey-admin-item" key={s.id}><div><h3>{s.title} <span className={s.is_active?'status active':'status inactive'}>{s.is_active?'فعال':'غیرفعال'}</span></h3><p>{s.description||'بدون توضیحات'}</p></div><div className="item-actions"><button className="secondary-button" onClick={async()=>{p.setSelected(s);await p.structure(s.id)}}> سؤال‌ها</button><button className="secondary-button" onClick={()=>{p.setSurveyForm({id:s.id,title:s.title,description:s.description||'',is_active:s.is_active});p.setModal('survey')}}>ویرایش</button><button className="danger-button" onClick={()=>p.removeSurvey(s.id)}>حذف</button></div></div>)}</div>{p.selected&&<div className="builder-panel"><div className="section-header"><div><h2>{p.selected.title}</h2><p>ساختار سؤال‌ها و گزینه‌ها</p></div><button className="secondary-button" onClick={()=>p.setSelected(null)}>بستن</button></div><div className="shared-options-panel">
<div className="shared-options-title">
  <div>
    <h3>⚡ گزینه‌های یکسان برای همه سؤال‌ها</h3>
    <p>
      گزینه‌هایی که اینجا تعریف می‌کنید، هنگام ساخت سؤال جدید
      به‌صورت خودکار اضافه می‌شوند.
    </p>
  </div>
</div>

<div className="shared-options-list">
  {p.sharedOptions.map((value:string,index:number)=>(
    <div className="shared-option-row" key={index}>

      <span className="shared-option-number">
        {index+1}
      </span>

      <input
        value={value}
        onChange={(e:any)=>
          p.setSharedOptions((items:string[])=>
            items.map((item:string,i:number)=>
              i===index ? e.target.value : item
            )
          )
        }
        placeholder={`گزینه ${index+1}`}
      />

      <button
        type="button"
        className="shared-option-delete"
        onClick={()=>
          p.setSharedOptions((items:string[])=>
            items.filter((_:string,i:number)=>i!==index)
          )
        }
      >
        🗑
      </button>

    </div>
  ))}
</div>

<div className="shared-options-actions">

  <button
    type="button"
    className="secondary-button"
    onClick={()=>
      p.setSharedOptions((items:string[])=>
        [...items,""]
      )
    }
  >
    ＋ افزودن گزینه
  </button>

  <button
    type="button"
    className="primary-button"
    disabled={p.busy}
    onClick={p.applySharedOptions}
  >
    ⚡ اعمال به همه سؤال‌ها
  </button>

</div>
</div>
<button className="primary-button" onClick={()=>{p.setQForm({id:0,text:''});p.setModal('question')}}>+ سؤال</button>{p.questions.filter((q:Question)=>q.survey_id===p.selected.id).map((q:Question,i:number)=><div className="question-card" key={q.id}><div className="question-head"><h3>{i+1}. {q.question_text}</h3><div className="item-actions"><button className="secondary-button" onClick={()=>{p.setQForm({id:q.id,text:q.question_text});p.setModal('question')}}>ویرایش</button><button className="danger-button" onClick={()=>p.removeQuestion(q.id)}>حذف</button></div></div>{p.options.filter((o:Option)=>o.question_id===q.id).map((o:Option)=><div className="option-row" key={o.id}>◉ {o.option_text}<span><button className="link-button" onClick={()=>{p.setOForm({id:o.id,question_id:q.id,text:o.option_text});p.setModal('option')}}>ویرایش</button><button className="link-button danger-text" onClick={()=>p.removeOption(o.id)}>حذف</button></span></div>)}<button className="add-option" onClick={()=>{p.setOForm({id:0,question_id:q.id,text:''});p.setModal('option')}}>+ افزودن گزینه</button></div>)}</div>}</section>}

function Users({p}:any){const filtered=p.users.filter((u:Profile)=>{const q=p.userSearch.trim().toLowerCase();return !q||u.username.toLowerCase().includes(q)||u.full_name.toLowerCase().includes(q)||(u.auth_email||"").toLowerCase().includes(q)});return <section className="content-card"><div className="section-header"><div><h2>کاربران</h2><p>افزودن، حذف، رمز عبور و سطح دسترسی</p></div><button className="primary-button" onClick={()=>{p.setUForm({id:'',username:'',full_name:'',email:'',password:'',is_admin:false,is_active:true,can_view_surveys:true,can_view_answers:true});p.setModal('user')}}>+ کاربر جدید</button></div><div className="toolbar"><input className="search-input" value={p.userSearch} onChange={(e:any)=>p.setUserSearch(e.target.value)} placeholder="جستجوی نام، نام کاربری یا ایمیل..."/><span className="toolbar-count">{filtered.length} کاربر</span></div><div className="table-wrap"><table><thead><tr><th>کاربر</th><th>ایمیل</th><th>نوع</th><th>وضعیت</th><th>نظرسنجی</th><th>پاسخ‌ها</th><th>عملیات</th></tr></thead><tbody>{filtered.map((u:Profile)=><tr key={u.id}><td><strong>{u.full_name}</strong><small>{u.username}</small></td><td>{u.auth_email||'—'}</td><td>{u.is_admin?'مدیر':'کاربر'}</td><td><span className={u.is_active?'status active':'status inactive'}>{u.is_active?'فعال':'غیرفعال'}</span></td><td><button className="toggle-button" onClick={()=>p.toggleUser(u,'can_view_surveys')}>{u.can_view_surveys!==false?'فعال':'بسته'}</button></td><td><button className="toggle-button" onClick={()=>p.toggleUser(u,'can_view_answers')}>{u.can_view_answers!==false?'فعال':'بسته'}</button></td><td><div className="item-actions"><button className="secondary-button" onClick={()=>{p.setUForm({id:u.id,username:u.username,full_name:u.full_name,email:u.auth_email||'',password:'',is_admin:u.is_admin,is_active:u.is_active,can_view_surveys:u.can_view_surveys!==false,can_view_answers:u.can_view_answers!==false});p.setModal('user')}}>ویرایش</button><button className="secondary-button" onClick={()=>p.toggleUser(u,'is_active')}>{u.is_active?'غیرفعال':'فعال'}</button></div></td></tr>)}</tbody></table></div></section>}

function Answers({p}:any){
 const filteredCompletions=p.completions.filter((c:any)=>!p.answerSurveyFilter||c.survey_id===p.answerSurveyFilter);
 const rows=p.users.filter((u:Profile)=>!u.is_admin||p.profile.can_view_answers!==false).map((u:Profile)=>({
  user:u,
  items:filteredCompletions.filter((c:any)=>c.user_id===u.id).map((c:any)=>({...c,survey:p.surveys.find((s:Survey)=>s.id===c.survey_id)}))
 }));
 return <section className="content-card">
  <div className="section-header"><div><h2>پاسخ‌های کاربران</h2><p>مشاهده مشارکت کاربران در نظرسنجی‌ها</p></div><div className="item-actions"><select className="filter-select" value={p.answerSurveyFilter} onChange={(e:any)=>p.setAnswerSurveyFilter(Number(e.target.value))}><option value={0}>همه نظرسنجی‌ها</option>{p.surveys.map((s:Survey)=><option key={s.id} value={s.id}>{s.title}</option>)}</select><button className="secondary-button" onClick={p.loadCompletions}>بروزرسانی</button></div></div>
  {!rows.some((r:any)=>r.items.length)?<div className="empty-state"><div>📝</div><h3>هنوز پاسخی ثبت نشده</h3></div>:
  <div className="answer-admin-list">{rows.filter((r:any)=>r.items.length).map((r:any)=><div className="answer-user-card" key={r.user.id}>
   <div className="answer-user-head"><div><strong>{r.user.full_name}</strong><small>{r.user.username}</small></div><span>{r.items.length} نظرسنجی</span></div>
   <div className="answer-survey-list">{r.items.map((x:any)=><div className="answer-survey-row" key={x.id}><div><strong>{x.survey?.title||`نظرسنجی #${x.survey_id}`}</strong><small>{new Date(x.created_at).toLocaleString('fa-IR')}</small></div><button className="secondary-button" onClick={async()=>{p.setViewUser(r.user);p.setViewSurvey(x.survey);await p.structure(x.survey_id);}}>مشاهده پاسخ‌ها</button></div>)}</div>
  </div>)}</div>}
  {p.viewUser&&p.viewSurvey&&<div className="modal-backdrop"><div className="modal large-modal"><div className="modal-header"><div><h2>{p.viewSurvey.title}</h2><p>پاسخ‌های {p.viewUser.full_name}</p></div><button className="close-button" onClick={()=>{p.setViewUser(null);p.setViewSurvey(null)}}>×</button></div>
   {p.questions.filter((q:Question)=>q.survey_id===p.viewSurvey.id).map((q:Question,i:number)=>{const vote=p.votes.find((v:Vote)=>v.user_id===p.viewUser.id&&v.question_id===q.id);const op=p.options.find((o:Option)=>o.id===vote?.option_id);return <div className="answer-question" key={q.id}><h3>{i+1}. {q.question_text}</h3><div className="answer-readonly">{op?.option_text||'پاسخی ثبت نشده'}</div></div>})}
   <div className="form-actions"><button className="secondary-button" onClick={()=>{p.setViewUser(null);p.setViewSurvey(null)}}>بستن</button></div>
  </div></div>}
 </section>
}

function Results({p}:any){const participants=p.selected?p.completions.filter((c:any)=>c.survey_id===p.selected.id).length:0;const eligible=p.users.filter((u:Profile)=>!u.is_admin&&u.is_active).length;const participation=eligible?Math.round(participants*100/eligible):0;return <section className="content-card"><div className="section-header"><div><h2>نتایج و نمودارها</h2><p>درصد و تعداد پاسخ هر گزینه</p></div><button className="secondary-button" onClick={p.loadVotes}>بروزرسانی</button></div><div className="result-summary">{p.selected&&<><div className="result-summary-card"><span>شرکت‌کنندگان</span><strong>{participants}</strong></div><div className="result-summary-card"><span>کاربران فعال</span><strong>{eligible}</strong></div><div className="result-summary-card"><span>درصد مشارکت</span><strong>{participation}%</strong></div></>}</div><div className="result-survey-picker">{p.surveys.map((s:Survey)=><button className={p.selected?.id===s.id?'survey-picker active':'survey-picker'} key={s.id} onClick={async()=>{p.setSelected(s);await p.structure(s.id);await p.loadVotes()}}>📋 {s.title}</button>)}</div>{!p.selected?<div className="empty-state"><div>📊</div><h3>یک نظرسنجی انتخاب کنید</h3></div>:<div className="results-list">{p.stats.map((st:any,i:number)=><div className="result-card" key={st.q.id}><h3>{i+1}. {st.q.question_text}</h3>{st.os.map((o:any)=><div className="chart-row" key={o.id}><div className="chart-label"><span>{o.option_text}</span><strong>{o.pct}% ({o.count})</strong></div><div className="bar-track"><div className="bar-fill" style={{width:`${o.pct}%`}}/></div></div>)}</div>)}</div>}</section>}

function Suggestions({p}:any){return <section className="content-card"><div className="section-header"><div><h2>پیشنهادهای کاربران</h2><p>یک پیشنهاد کلی برای هر نظرسنجی؛ اختیاری</p></div><button className="secondary-button" onClick={p.loadSuggestions}>بروزرسانی</button></div><div className="suggestion-admin-list">{!p.suggestions.length?<div className="empty-state"><div>💬</div><h3>پیشنهادی ثبت نشده</h3></div>:p.suggestions.map((s:Suggestion)=><div className="suggestion-admin-item" key={s.id}><strong>{s.profiles?.full_name||s.profiles?.username||'کاربر'}</strong><small>{p.surveys.find((x:Survey)=>x.id===s.survey_id)?.title||`نظرسنجی #${s.survey_id}`}</small><p>{s.suggestion_text}</p><small>{new Date(s.created_at).toLocaleString('fa-IR')}</small></div>)}</div></section>}

function Modals({p}:any){return <>{p.modal==='survey'&&<Modal title={p.surveyForm.id?'ویرایش نظرسنجی':'نظرسنجی جدید'} close={()=>p.setModal(null)}><form onSubmit={p.saveSurvey}><label>عنوان</label><input value={p.surveyForm.title} onChange={(e:any)=>p.setSurveyForm({...p.surveyForm,title:e.target.value})}/><label>توضیحات</label><textarea value={p.surveyForm.description} onChange={(e:any)=>p.setSurveyForm({...p.surveyForm,description:e.target.value})}/><label className="checkbox-row"><input type="checkbox" checked={p.surveyForm.is_active} onChange={(e:any)=>p.setSurveyForm({...p.surveyForm,is_active:e.target.checked})}/> فعال</label><Actions close={()=>p.setModal(null)}/></form></Modal>}{p.modal==='question'&&<Modal title="سؤال" close={()=>p.setModal(null)}><form onSubmit={p.saveQuestion}><label>متن سؤال</label><input value={p.qForm.text} onChange={(e:any)=>p.setQForm({...p.qForm,text:e.target.value})}/><Actions close={()=>p.setModal(null)}/></form></Modal>}{p.modal==='option'&&<Modal title="گزینه" close={()=>p.setModal(null)}><form onSubmit={p.saveOption}><label>متن گزینه</label><input value={p.oForm.text} onChange={(e:any)=>p.setOForm({...p.oForm,text:e.target.value})}/><Actions close={()=>p.setModal(null)}/></form></Modal>}{p.modal==='user'&&<Modal title={p.uForm.id?'ویرایش کاربر':'ایجاد کاربر'} close={()=>p.setModal(null)}><form onSubmit={p.saveUser}><label>نام کاربری</label><input value={p.uForm.username} onChange={(e:any)=>p.setUForm({...p.uForm,username:e.target.value})}/><label>نام کامل</label><input value={p.uForm.full_name} onChange={(e:any)=>p.setUForm({...p.uForm,full_name:e.target.value})}/>{!p.uForm.id&&<><label>ایمیل</label><input type="email" value={p.uForm.email} onChange={(e:any)=>p.setUForm({...p.uForm,email:e.target.value})}/><label>رمز عبور</label><input type="password" value={p.uForm.password} onChange={(e:any)=>p.setUForm({...p.uForm,password:e.target.value})}/></>}<label className="checkbox-row"><input type="checkbox" checked={p.uForm.is_admin} onChange={(e:any)=>p.setUForm({...p.uForm,is_admin:e.target.checked})}/> مدیر</label><label className="checkbox-row"><input type="checkbox" checked={p.uForm.is_active} onChange={(e:any)=>p.setUForm({...p.uForm,is_active:e.target.checked})}/> فعال</label><label className="checkbox-row"><input type="checkbox" checked={p.uForm.can_view_surveys} onChange={(e:any)=>p.setUForm({...p.uForm,can_view_surveys:e.target.checked})}/> مشاهده نظرسنجی‌ها</label><label className="checkbox-row"><input type="checkbox" checked={p.uForm.can_view_answers} onChange={(e:any)=>p.setUForm({...p.uForm,can_view_answers:e.target.checked})}/> مشاهده پاسخ‌ها</label><Actions close={()=>p.setModal(null)} busy={p.busy}/></form></Modal>}</>}
function Modal({title,close,children}:any){return <div className="modal-backdrop"><div className="modal"><div className="modal-header"><h2>{title}</h2><button className="close-button" onClick={close}>×</button></div>{children}</div></div>}
function Actions({close,busy}:any){return <div className="form-actions"><button type="button" className="secondary-button" onClick={close}>انصراف</button><button className="primary-button" disabled={busy}>{busy?'در حال ذخیره...':'ذخیره'}</button></div>}
