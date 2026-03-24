import { useState, useEffect } from 'react'
import { addSurveyResponse } from '../lib/storage.js'

const QUESTIONS = [
  { key: 'happiness',  label: 'How is your overall happiness at work right now?' },
  { key: 'teamSpirit', label: 'How would you rate the team spirit in your team?' },
  { key: 'belonging',  label: 'How strong is your sense of belonging in your team?' },
]

function StarRating({ value, onChange }) {
  const [hover, setHover] = useState(0)
  return (
    <div className="star-row" role="group" aria-label="Rating 1 to 5">
      {[1, 2, 3, 4, 5].map(n => (
        <button
          key={n}
          className={`star-btn ${n <= (hover || value) ? 'active' : ''}`}
          onClick={() => onChange(n)}
          onMouseEnter={() => setHover(n)}
          onMouseLeave={() => setHover(0)}
          aria-label={`${n} star${n > 1 ? 's' : ''}`}
          type="button"
        >
          {n <= (hover || value) ? '★' : '☆'}
        </button>
      ))}
    </div>
  )
}

export default function PulseSurvey() {
  const params    = new URLSearchParams(window.location.search)
  const teamName  = params.get('team') || 'Your team'
  const plName    = params.get('pl')   || ''

  const [scores,    setScores]    = useState({ happiness: 0, teamSpirit: 0, belonging: 0 })
  const [submitted, setSubmitted] = useState(false)
  const [count,     setCount]     = useState(null)
  const [error,     setError]     = useState(null)

  const allAnswered = scores.happiness > 0 && scores.teamSpirit > 0 && scores.belonging > 0

  function setScore(key, val) {
    setScores(s => ({ ...s, [key]: val }))
  }

  function handleSubmit() {
    if (!allAnswered) return
    try {
      const total = addSurveyResponse(plName, teamName, scores)
      setCount(total)
      setSubmitted(true)
    } catch (e) {
      setError('Something went wrong. Please try again.')
    }
  }

  if (submitted) {
    return (
      <div className="page-survey">
        <div className="survey-wrap">
          <div className="survey-thankyou">
            <div style={{ fontSize: 64, marginBottom: 16 }}>★</div>
            <h2>Thank you!</h2>
            <p>Your response has been recorded anonymously.</p>
            {count !== null && count < 5 && (
              <p style={{ marginTop: 12, fontSize: 13, opacity: .6 }}>
                {5 - count} more response{5 - count !== 1 ? 's' : ''} needed before scores are visible.
              </p>
            )}
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="page-survey">
      <div className="survey-wrap">
        {/* Header */}
        <div className="survey-header">
          <p style={{ fontSize: 12, color: 'var(--accent)', fontWeight: 700, letterSpacing: '.08em', textTransform: 'uppercase', marginBottom: 8 }}>
            PL Report Tool  ·  Pulse Survey
          </p>
          <h1>{teamName}</h1>
          <p>Anonymous · 3 questions · Takes 30 seconds</p>
        </div>

        {/* Questions */}
        <div className="survey-body">
          {error && (
            <div className="error-bar" style={{ marginBottom: 14 }}>{error}</div>
          )}

          {QUESTIONS.map(q => (
            <div key={q.key} className="survey-question">
              <h2>{q.label}</h2>
              <StarRating
                value={scores[q.key]}
                onChange={val => setScore(q.key, val)}
              />
              {scores[q.key] > 0 && (
                <p style={{ textAlign: 'center', marginTop: 12, fontSize: 13, color: 'var(--text-muted)' }}>
                  {['', 'Not great', 'Below average', 'Okay', 'Good', 'Excellent'][scores[q.key]]}
                </p>
              )}
            </div>
          ))}

          <button
            className="survey-submit"
            onClick={handleSubmit}
            disabled={!allAnswered}
          >
            {allAnswered ? 'Submit anonymously' : 'Please answer all 3 questions'}
          </button>

          <p style={{ textAlign: 'center', marginTop: 14, fontSize: 12, color: 'rgba(255,255,255,.4)' }}>
            Responses are anonymous. Individual scores are never shown.
            Results only visible once 5+ team members respond.
          </p>
        </div>
      </div>
    </div>
  )
}
