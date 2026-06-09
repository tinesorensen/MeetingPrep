import dynamic from 'next/dynamic'

const MeetingPrep = dynamic(() => import('../components/MeetingPrep'), { ssr: false })

export default function Home() {
  return <MeetingPrep />
}
