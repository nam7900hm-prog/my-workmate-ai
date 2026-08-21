import './style.css';
import './student.css';
import './modal.css';
import './analysis.css';
export const metadata={title:'my workmate ai',description:'자료를 넣고 말하면 결과를 만드는 개인 업무비서'};
export default function Layout({children}:{children:React.ReactNode}){return <html lang="ko"><body>{children}</body></html>}
