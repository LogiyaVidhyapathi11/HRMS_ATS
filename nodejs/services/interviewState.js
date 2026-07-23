/**
 * In-memory store for active interviews.
 * Key: threadId or callId
 * Value: {
 *   candidateName: string,
 *   questions: Array<{ text: string, type: string }>,
 *   audioUrls: string[],  // Public URLs for the TTS audio of each question
 *   currentIndex: number, // Which question is currently being asked
 *   status: string        // "waiting", "asking", "listening", "completed"
 *   meetingId: string     // Graph onlineMeeting ID - used to fetch full session recording.
 * }
 */
const activeInterviews = new Map();

function createInterviewState(threadId, candidateName, candidateEmail, questions, audioUrls, endTime, meetingId, organizerId, eventId) {
    activeInterviews.set(threadId, {
        threadId,
        candidateName, 
        candidateEmail, 
        questions,
        audioUrls,
        endTime,
        meetingId: meetingId || null, 
        organizerId: organizerId || null, 
        eventId: eventId || null, 
        consecutiveTimeouts: 0, 
        currentIndex: 0,
        status: "waiting",
        hasHumanJoined: false, 
        interviewStarted: false, // Prevents duplicate interview loops
        warningsSent: {
            min5: false, 
            min1: false, 
            timesUp: false
        }
    });
}

function getInterviewState(threadId) {
    return activeInterviews.get(threadId);
}

function getInterviewStateByCallId(callId) {
    for (const [threadId, state] of activeInterviews.entries()) {
        if (state.callId === callId) {
            return state;
        }
    }
    return null;
}

function setCallIdForThread(threadId, callId) {
    const state = activeInterviews.get(threadId);
    if (state) {
        state.callId = callId;
    }
}

function incrementQuestion(callId) {
    const state = getInterviewStateByCallId(callId);
    if (state) {
        state.currentIndex++;
        return state.currentIndex;
    }
    return -1;
}

function deleteInterviewState(threadId) {
    activeInterviews.delete(threadId);
}

module.exports = {
    createInterviewState,
    getInterviewState,
    setCallIdForThread,
    getInterviewStateByCallId,
    incrementQuestion,
    deleteInterviewState
};
